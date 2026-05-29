import { FormEvent, KeyboardEvent, useMemo, useRef, useState } from "react";
import { runAgentLoop, type AgentMessage as PiAgentMessage } from "@earendil-works/pi-agent-core/browser";
import { getModels } from "@earendil-works/pi-ai/models";
import { streamSimpleOpenAIResponses } from "@earendil-works/pi-ai/openai-responses";
import type { AssistantMessage, Message, Model, TextContent, ToolCall, UserMessage } from "@earendil-works/pi-ai/types";
import { createExecuteOfficeJsCodeTool, EXECUTE_OFFICEJS_CODE_TOOL_NAME } from "@milton/office-runtime/tool";

export interface ExcelTaskpaneAppProps {
  /** OpenAI credentials and model selection used by the taskpane agent loop. */
  openAI: {
    /** API key used for taskpane-local model requests. */
    apiKey: string;
    /** OpenAI model id selected for Milton. */
    model: string;
  };
  /** Optional development profile label displayed in debug builds. */
  devProfile?: {
    /** Human-readable profile name for the active taskpane configuration. */
    label: string;
  };
}

const MILTON_SYSTEM_PROMPT = [
  "You are Milton, an AI assistant running in an Excel task pane.",
  "Use execute_officejs_code whenever you need to inspect, read, edit, or automate the current workbook.",
  "The tool compiles and runs TypeScript against Excel through the Excel OfficeJS API.",
  "Write code as export async function run(ctx: ExcelRuntimeContext), use ctx.workbook for workbook objects, and use ctx.context when an OfficeJS API requires the raw Excel.RequestContext.",
  "Batch OfficeJS load and mutation calls before awaiting ctx.sync(); after load() calls, await ctx.sync() before reading loaded properties.",
  "Return a JSON-serializable object from run(ctx) for any workbook data, status, or results you need to see after the tool runs.",
  "Do not import packages, access the DOM, make network requests, or use browser globals from generated workbook code.",
  "The execution tool is not sandboxed in this milestone, so only run code that is directly relevant to the user's workbook request.",
  "Be concise, practical, and clear.",
].join("\n");

type TaskpaneMessageStatus = "complete" | "streaming" | "error" | "cancelled";

interface TaskpaneMessage {
  /** Stable UI message id. */
  id: string;
  /** Message author or tool source. */
  role: "user" | "assistant" | "tool";
  /** Primary visible message text. */
  content: string;
  /** ISO timestamp shown for ordering and reset stability. */
  createdAt: string;
  /** Current lifecycle status for the message. */
  status: TaskpaneMessageStatus;
  /** Optional workbook tool debug details. */
  toolDetails?: TaskpaneToolDetails;
}

interface TaskpaneToolDetails {
  /** Tool display name. */
  toolName: string;
  /** TypeScript source sent to the OfficeJS code tool. */
  code?: string;
  /** Logs emitted by generated code. */
  logs: TaskpaneToolLog[];
  /** Compile diagnostics returned by the OfficeJS compiler. */
  diagnostics: TaskpaneDiagnostic[];
  /** JSON-compatible value returned by generated code. */
  returnValue?: unknown;
  /** Elapsed execution time in milliseconds. */
  elapsedMs?: number;
}

interface TaskpaneToolLog {
  /** Log message emitted from generated code. */
  message: string;
  /** Optional structured log details. */
  details?: unknown;
  /** Unix timestamp recorded by the runtime. */
  timestamp?: number;
}

interface TaskpaneDiagnostic {
  /** Diagnostic severity label. */
  severity: string;
  /** Diagnostic message text. */
  message: string;
  /** One-based source line when available. */
  line?: number;
  /** One-based source column when available. */
  column?: number;
  /** TypeScript or Milton-specific diagnostic code. */
  code?: number | string;
}

const INITIAL_MESSAGES: TaskpaneMessage[] = [
  {
    id: "welcome",
    role: "assistant",
    content: "I am Milton. Ask me to inspect or edit this workbook.",
    createdAt: new Date().toISOString(),
    status: "complete",
  },
];

/** Renders the Excel taskpane chat surface and wires Milton to workbook tools. */
export function ExcelTaskpaneApp({ devProfile, openAI }: ExcelTaskpaneAppProps) {
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<TaskpaneMessage[]>(INITIAL_MESSAGES);
  const [isRunning, setIsRunning] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const messagesRef = useRef(messages);
  const piMessagesRef = useRef<PiAgentMessage[]>([]);
  const toolMessageIdsRef = useRef(new Map<string, string>());
  const isConfigured = openAI.apiKey.trim().length > 0;
  const model = useMemo(() => resolveOpenAIResponsesModel(openAI.model), [openAI.model]);
  const tools = useMemo(() => [createExecuteOfficeJsCodeTool()], []);

  /** Replaces taskpane messages while keeping the mutable event-handler ref in sync. */
  function updateMessages(nextMessages: TaskpaneMessage[]) {
    messagesRef.current = nextMessages;
    setMessages(nextMessages);
  }

  /** Appends one taskpane message. */
  function appendMessage(message: TaskpaneMessage) {
    updateMessages([...messagesRef.current, message]);
  }

  /** Applies a shallow patch to one taskpane message. */
  function updateMessage(id: string, patch: Partial<TaskpaneMessage>) {
    updateMessages(messagesRef.current.map((message) => (message.id === id ? { ...message, ...patch } : message)));
  }

  /** Updates one taskpane message with access to its existing value. */
  function updateMessageWith(id: string, updater: (message: TaskpaneMessage) => TaskpaneMessage) {
    updateMessages(messagesRef.current.map((message) => (message.id === id ? updater(message) : message)));
  }

  /** Sends the current user prompt through the Pi agent loop. */
  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    const userContent = draft.trim();

    if (!userContent || isRunning) {
      return;
    }

    setDraft("");
    setIsRunning(true);
    setStatus(null);

    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    try {
      const userMessage: UserMessage = {
        role: "user",
        content: [{ type: "text", text: userContent }],
        timestamp: Date.now(),
      };
      const userMessageId = createTaskpaneMessageId("user");
      let assistantMessageId: string | undefined;

      const nextPiMessages = await runAgentLoop(
        [userMessage],
        {
          systemPrompt: MILTON_SYSTEM_PROMPT,
          messages: piMessagesRef.current,
          tools,
        },
        {
          apiKey: openAI.apiKey,
          cacheRetention: "none",
          convertToLlm: (agentMessages) => agentMessages.filter(isLlmMessage),
          model,
          shouldStopAfterTurn: ({ toolResults }) => toolResults.length === 0,
          toolExecution: "sequential",
        },
        (agentEvent) => {
          if (agentEvent.type === "tool_execution_start") {
            const toolMessageId = createTaskpaneMessageId("tool");
            toolMessageIdsRef.current.set(agentEvent.toolCallId, toolMessageId);
            appendMessage({
              id: toolMessageId,
              role: "tool",
              content: getToolExecutionStartText(agentEvent.toolName),
              createdAt: new Date().toISOString(),
              status: "streaming",
              toolDetails: {
                toolName: agentEvent.toolName,
                code: getToolCode(agentEvent.args),
                logs: [],
                diagnostics: [],
              },
            });
            setStatus(getToolExecutionStartText(agentEvent.toolName));
            return;
          }

          if (agentEvent.type === "tool_execution_update") {
            const toolMessageId = toolMessageIdsRef.current.get(agentEvent.toolCallId);

            if (toolMessageId) {
              appendToolUpdate(toolMessageId, agentEvent.partialResult);
            }

            return;
          }

          if (agentEvent.type === "tool_execution_end") {
            const toolMessageId = toolMessageIdsRef.current.get(agentEvent.toolCallId);

            if (toolMessageId) {
              updateToolMessageFromResult(toolMessageId, agentEvent.toolName, agentEvent.result, agentEvent.isError);
            }

            toolMessageIdsRef.current.delete(agentEvent.toolCallId);
            setStatus(agentEvent.isError ? getToolExecutionErrorText(agentEvent.result) : null);
            return;
          }

          if (agentEvent.type === "message_start" && agentEvent.message.role === "user") {
            appendMessage({
              id: userMessageId,
              role: "user",
              content: getMessageText(agentEvent.message),
              createdAt: new Date(agentEvent.message.timestamp).toISOString(),
              status: "complete",
            });
            return;
          }

          if (agentEvent.type === "message_start" && agentEvent.message.role === "assistant") {
            assistantMessageId = createTaskpaneMessageId("assistant");
            appendMessage({
              id: assistantMessageId,
              role: "assistant",
              content: getMessageText(agentEvent.message),
              createdAt: new Date(agentEvent.message.timestamp).toISOString(),
              status: "streaming",
            });
            return;
          }

          if (agentEvent.type === "message_update" && agentEvent.message.role === "assistant" && assistantMessageId) {
            updateMessage(assistantMessageId, {
              content: getMessageText(agentEvent.message),
              status: "streaming",
            });
            return;
          }

          if (agentEvent.type === "message_end" && agentEvent.message.role === "assistant") {
            const id = assistantMessageId ?? createTaskpaneMessageId("assistant");
            const status = getAssistantStatus(agentEvent.message);

            if (!assistantMessageId) {
              assistantMessageId = id;
              appendMessage({
                id,
                role: "assistant",
                content: getMessageText(agentEvent.message),
                createdAt: new Date(agentEvent.message.timestamp).toISOString(),
                status,
              });
              return;
            }

            updateMessage(id, {
              content: getMessageText(agentEvent.message),
              status,
            });

            if (status === "error") {
              setStatus(agentEvent.message.errorMessage ?? "The agent hit an unknown error.");
            }
          }
        },
        abortController.signal,
        streamSimpleOpenAIResponses,
      );

      piMessagesRef.current = [...piMessagesRef.current, ...nextPiMessages];
    } catch (error) {
      console.error(error);
      setStatus(error instanceof Error ? error.message : "The agent hit an unknown error.");
    } finally {
      abortControllerRef.current = null;
      setIsRunning(false);
    }
  }

  /** Appends a streamed OfficeJS log update to an existing tool message. */
  function appendToolUpdate(messageId: string, partialResult: unknown) {
    const latestLog = getLatestToolLog(partialResult);

    if (!latestLog) {
      return;
    }

    updateMessageWith(messageId, (message) => {
      const previousDetails = message.toolDetails ?? createEmptyToolDetails("execute_officejs_code");

      return {
        ...message,
        content: `OfficeJS log: ${latestLog.message}`,
        toolDetails: {
          ...previousDetails,
          code: previousDetails.code ?? getToolCodeFromPartial(partialResult),
          logs: previousDetails.logs.concat(latestLog),
        },
      };
    });
  }

  /** Updates a tool message with final execution details. */
  function updateToolMessageFromResult(messageId: string, toolName: string, result: unknown, isError: boolean) {
    const executionDetails = getExecutionDetails(result);

    updateMessageWith(messageId, (message) => {
      const previousDetails = message.toolDetails ?? createEmptyToolDetails(toolName);
      const nextDetails = executionDetails
        ? {
            ...previousDetails,
            code: executionDetails.code ?? previousDetails.code,
            logs: executionDetails.logs ?? previousDetails.logs,
            diagnostics: executionDetails.diagnostics ?? previousDetails.diagnostics,
            returnValue: executionDetails.returnValue,
            elapsedMs: executionDetails.elapsedMs,
          }
        : previousDetails;

      return {
        ...message,
        content: isError ? getToolExecutionErrorText(result) : getToolExecutionSuccessText(nextDetails),
        status: isError ? "error" : "complete",
        toolDetails: nextDetails,
      };
    });
  }

  /** Cancels the active taskpane agent run. */
  function handleCancel() {
    abortControllerRef.current?.abort();
  }

  /** Clears taskpane transcript state and cancels active work if needed. */
  function handleReset() {
    if (isRunning) {
      abortControllerRef.current?.abort();
    }

    updateMessages(INITIAL_MESSAGES);
    piMessagesRef.current = [];
    toolMessageIdsRef.current.clear();
    setDraft("");
    setStatus(null);
  }

  /** Submits the composer on Enter while preserving Shift+Enter newlines. */
  function handleComposerKeyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      event.currentTarget.form?.requestSubmit();
    }
  }

  return (
    <main className="taskpane-shell">
      <header className="taskpane-header">
        <div className="brand-lockup">
          <img className="logo" src="/assets/logo-filled.png" alt="" />
          <div>
            <h1>Milton</h1>
            <p>{openAI.model}</p>
          </div>
        </div>
        <button className="secondary-button" type="button" onClick={handleReset}>
          Reset
        </button>
      </header>
      {devProfile?.label ? <p className="dev-profile-label">{devProfile.label}</p> : null}

      <section className="conversation" aria-label="Conversation">
        {messages.map((message) => (
          <article className={`message message-${message.role}`} key={message.id}>
            <div className="message-meta">
              <span>{getMessageRoleLabel(message.role)}</span>
              {message.status === "streaming" ? <span>Thinking</span> : null}
              {message.status === "error" ? <span>Error</span> : null}
              {message.status === "cancelled" ? <span>Cancelled</span> : null}
            </div>
            <p>{message.content || " "}</p>
            {message.toolDetails ? <ToolDetailsView details={message.toolDetails} /> : null}
          </article>
        ))}
      </section>

      <form className="composer" onSubmit={handleSubmit}>
        {!isConfigured ? <p className="configuration-warning">Missing DEBUG_OPENAI_API_KEY in apps/office/.env.local.</p> : null}
        {status ? <p className="status">{status}</p> : null}
        <label className="sr-only" htmlFor="message-draft">
          Message
        </label>
        <textarea
          id="message-draft"
          value={draft}
          onChange={(event) => setDraft(event.target.value)}
          onKeyDown={handleComposerKeyDown}
          placeholder="Message Milton"
          rows={3}
        />
        <div className="composer-actions">
          <button className="secondary-button" type="button" onClick={handleCancel} disabled={!isRunning}>
            Stop
          </button>
          <button className="primary-button" type="submit" disabled={!draft.trim() || isRunning || !isConfigured}>
            Send
          </button>
        </div>
      </form>
    </main>
  );
}

/** Renders expandable OfficeJS tool execution details. */
function ToolDetailsView({ details }: { details: TaskpaneToolDetails }) {
  const logsText = details.logs.length > 0 ? formatJson(details.logs) : "";
  const diagnosticsText = details.diagnostics.length > 0 ? formatJson(details.diagnostics) : "";
  const returnValueText = details.returnValue === undefined ? "" : formatJson(details.returnValue);

  return (
    <div className="tool-details">
      {details.elapsedMs !== undefined ? <p className="tool-elapsed">{Math.round(details.elapsedMs)} ms</p> : null}
      {details.code ? (
        <details>
          <summary>Code</summary>
          <pre>{details.code}</pre>
        </details>
      ) : null}
      {logsText ? (
        <details>
          <summary>Logs</summary>
          <pre>{logsText}</pre>
        </details>
      ) : null}
      {diagnosticsText ? (
        <details>
          <summary>Diagnostics</summary>
          <pre>{diagnosticsText}</pre>
        </details>
      ) : null}
      {returnValueText ? (
        <details>
          <summary>Returned data</summary>
          <pre>{returnValueText}</pre>
        </details>
      ) : null}
    </div>
  );
}

/** Builds an empty tool details object for streamed updates. */
function createEmptyToolDetails(toolName: string): TaskpaneToolDetails {
  return {
    toolName,
    logs: [],
    diagnostics: [],
  };
}

/** Creates a taskpane message id with a role prefix for easier debugging. */
function createTaskpaneMessageId(role: TaskpaneMessage["role"]): string {
  return `${role}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/** Returns the visible label for a taskpane message role. */
function getMessageRoleLabel(role: TaskpaneMessage["role"]): string {
  if (role === "assistant") {
    return "Milton";
  }

  if (role === "tool") {
    return "Workbook Tool";
  }

  return "You";
}

/** Keeps only Pi messages that the configured model provider accepts. */
function isLlmMessage(message: PiAgentMessage): message is Message {
  return message.role === "user" || message.role === "assistant" || message.role === "toolResult";
}

/** Maps assistant stop reasons to taskpane message status values. */
function getAssistantStatus(message: AssistantMessage): TaskpaneMessageStatus {
  if (message.stopReason === "aborted") {
    return "cancelled";
  }

  if (message.stopReason === "error") {
    return "error";
  }

  return "complete";
}

/** Extracts visible text from Pi agent messages for the taskpane transcript. */
function getMessageText(message: PiAgentMessage): string {
  if (message.role === "user") {
    if (typeof message.content === "string") {
      return message.content;
    }

    return message.content
      .map((content) => (content.type === "text" ? content.text : "[image]"))
      .join("\n")
      .trim();
  }

  if (message.role === "assistant") {
    const text = message.content
      .filter((content): content is TextContent => content.type === "text")
      .map((content) => content.text)
      .join("\n")
      .trim();

    if (text) {
      return text;
    }

    const toolCallText = getAssistantToolCallText(message);

    if (toolCallText) {
      return toolCallText;
    }

    if (message.errorMessage) {
      return message.errorMessage;
    }

    return "";
  }

  return "";
}

/** Builds visible text for assistant messages that only contain tool calls. */
function getAssistantToolCallText(message: AssistantMessage): string {
  const toolCalls = message.content.filter((content): content is ToolCall => content.type === "toolCall");

  if (toolCalls.some((toolCall) => toolCall.name === EXECUTE_OFFICEJS_CODE_TOOL_NAME)) {
    return "Using the Excel OfficeJS workbook tool.";
  }

  if (toolCalls.length > 0) {
    return "Using a tool.";
  }

  return "";
}

/** Builds taskpane status text when a tool starts running. */
function getToolExecutionStartText(toolName: string): string {
  if (toolName === EXECUTE_OFFICEJS_CODE_TOOL_NAME) {
    return "Running OfficeJS workbook code.";
  }

  return `Running ${toolName}.`;
}

/** Builds taskpane text for a completed workbook tool result. */
function getToolExecutionSuccessText(details: TaskpaneToolDetails): string {
  if (details.returnValue !== undefined) {
    return "OfficeJS workbook code finished with returned data.";
  }

  if (details.logs.length > 0) {
    return "OfficeJS workbook code finished with logs.";
  }

  return "OfficeJS workbook code finished.";
}

/** Extracts readable taskpane status text from an errored tool result. */
function getToolExecutionErrorText(result: unknown): string {
  const content = (result as { content?: Array<{ text?: unknown; type?: unknown }> }).content;
  const text =
    content
      ?.filter((block) => block.type === "text" && typeof block.text === "string")
      .map((block) => block.text)
      .join("\n")
      .trim() ?? "";

  return text || "OfficeJS workbook code failed.";
}

/** Extracts the code argument from a tool start event payload. */
function getToolCode(args: unknown): string | undefined {
  const record = asRecord(args);
  return typeof record?.code === "string" ? record.code : undefined;
}

/** Extracts streamed code from a partial tool result. */
function getToolCodeFromPartial(partialResult: unknown): string | undefined {
  const details = asRecord(asRecord(partialResult)?.details);
  return typeof details?.code === "string" ? details.code : undefined;
}

/** Extracts a streamed OfficeJS log entry from a partial tool result. */
function getLatestToolLog(partialResult: unknown): TaskpaneToolLog | undefined {
  const details = asRecord(asRecord(partialResult)?.details);
  const latestLog = asRecord(details?.latestLog);

  if (typeof latestLog?.message !== "string") {
    return undefined;
  }

  return {
    message: latestLog.message,
    details: latestLog.details,
    timestamp: typeof latestLog.timestamp === "number" ? latestLog.timestamp : undefined,
  };
}

/** Extracts final OfficeJS execution details from a tool result payload. */
function getExecutionDetails(result: unknown): Partial<TaskpaneToolDetails> | undefined {
  const details = asRecord(asRecord(result)?.details);

  if (!details) {
    return undefined;
  }

  return {
    code: typeof details.code === "string" ? details.code : undefined,
    logs: parseToolLogs(details.logs),
    diagnostics: parseDiagnostics(details.diagnostics),
    returnValue: details.returnValue,
    elapsedMs: typeof details.elapsedMs === "number" ? details.elapsedMs : undefined,
  };
}

/** Converts an unknown logs payload into taskpane log entries. */
function parseToolLogs(value: unknown): TaskpaneToolLog[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  return value.flatMap((entry) => {
    const record = asRecord(entry);

    if (typeof record?.message !== "string") {
      return [];
    }

    return [
      {
        message: record.message,
        details: record.details,
        timestamp: typeof record.timestamp === "number" ? record.timestamp : undefined,
      },
    ];
  });
}

/** Converts an unknown diagnostics payload into taskpane diagnostics. */
function parseDiagnostics(value: unknown): TaskpaneDiagnostic[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  return value.flatMap((entry) => {
    const record = asRecord(entry);

    if (typeof record?.message !== "string" || typeof record.severity !== "string") {
      return [];
    }

    return [
      {
        severity: record.severity,
        message: record.message,
        line: typeof record.line === "number" ? record.line : undefined,
        column: typeof record.column === "number" ? record.column : undefined,
        code: typeof record.code === "number" || typeof record.code === "string" ? record.code : undefined,
      },
    ];
  });
}

/** Safely formats JSON-compatible values for taskpane detail panels. */
function formatJson(value: unknown): string {
  try {
    return JSON.stringify(value, null, 2) ?? "";
  } catch {
    return String(value);
  }
}

/** Narrows unknown values to records for payload extraction. */
function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null ? (value as Record<string, unknown>) : undefined;
}

/** Resolves a configured model id into the Pi OpenAI Responses model shape. */
function resolveOpenAIResponsesModel(modelId: string): Model<"openai-responses"> {
  const registryModel = getModels("openai").find((candidate) => candidate.id === modelId);

  if (registryModel?.api === "openai-responses") {
    return registryModel as Model<"openai-responses">;
  }

  return {
    id: modelId,
    name: modelId,
    api: "openai-responses",
    provider: "openai",
    baseUrl: "https://api.openai.com/v1",
    reasoning: false,
    input: ["text", "image"],
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    },
    contextWindow: 128000,
    maxTokens: 8192,
  };
}
