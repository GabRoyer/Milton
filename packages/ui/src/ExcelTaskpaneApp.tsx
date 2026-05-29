import { FormEvent, KeyboardEvent, useMemo, useRef, useState } from "react";
import { runAgentLoop, type AgentMessage as PiAgentMessage } from "@earendil-works/pi-agent-core/browser";
import { getModels } from "@earendil-works/pi-ai/models";
import { streamSimpleOpenAIResponses } from "@earendil-works/pi-ai/openai-responses";
import type { AssistantMessage, Message, Model, TextContent, ToolCall, UserMessage } from "@earendil-works/pi-ai/types";
import { createExecuteOfficeJsCodeTool, EXECUTE_OFFICEJS_CODE_TOOL_NAME } from "@milton/office-runtime/tool";

export interface ExcelTaskpaneAppProps {
  openAI: {
    apiKey: string;
    model: string;
  };
  devProfile?: {
    label: string;
  };
}

const MILTON_SYSTEM_PROMPT = [
  "You are Milton, an AI assistant running in an Excel task pane.",
  "Use execute_officejs_code whenever you need to inspect, read, edit, or automate the current workbook.",
  "The tool compiles and runs TypeScript against Excel through the Excel OfficeJS API.",
  "Write code as an async run(ctx: ExcelRuntimeContext) function and use ctx.context, ctx.workbook, and ctx.sync() for OfficeJS batching.",
  "Return a JSON-serializable object from run(ctx) for any workbook data or results you need to see after the tool runs.",
  "Do not import packages, access the DOM, make network requests, or use browser globals from generated workbook code.",
  "The execution tool is not sandboxed in this milestone, so only run code that is directly relevant to the user's workbook request.",
  "Be concise, practical, and clear.",
].join("\n");

type TaskpaneMessageStatus = "complete" | "streaming" | "error" | "cancelled";

interface TaskpaneMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  createdAt: string;
  status: TaskpaneMessageStatus;
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

export function ExcelTaskpaneApp({ devProfile, openAI }: ExcelTaskpaneAppProps) {
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<TaskpaneMessage[]>(INITIAL_MESSAGES);
  const [isRunning, setIsRunning] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const messagesRef = useRef(messages);
  const piMessagesRef = useRef<PiAgentMessage[]>([]);
  const isConfigured = openAI.apiKey.trim().length > 0;
  const model = useMemo(() => resolveOpenAIResponsesModel(openAI.model), [openAI.model]);
  const tools = useMemo(() => [createExecuteOfficeJsCodeTool()], []);

  function updateMessages(nextMessages: TaskpaneMessage[]) {
    messagesRef.current = nextMessages;
    setMessages(nextMessages);
  }

  function appendMessage(message: TaskpaneMessage) {
    updateMessages([...messagesRef.current, message]);
  }

  function updateMessage(id: string, patch: Partial<TaskpaneMessage>) {
    updateMessages(messagesRef.current.map((message) => (message.id === id ? { ...message, ...patch } : message)));
  }

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
            setStatus(getToolExecutionStartText(agentEvent.toolName));
            return;
          }

          if (agentEvent.type === "tool_execution_end") {
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

  function handleCancel() {
    abortControllerRef.current?.abort();
  }

  function handleReset() {
    if (isRunning) {
      abortControllerRef.current?.abort();
    }

    updateMessages(INITIAL_MESSAGES);
    piMessagesRef.current = [];
    setDraft("");
    setStatus(null);
  }

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
              <span>{message.role === "assistant" ? "Milton" : "You"}</span>
              {message.status === "streaming" ? <span>Thinking</span> : null}
              {message.status === "error" ? <span>Error</span> : null}
              {message.status === "cancelled" ? <span>Cancelled</span> : null}
            </div>
            <p>{message.content || " "}</p>
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

function createTaskpaneMessageId(role: TaskpaneMessage["role"]): string {
  return `${role}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function isLlmMessage(message: PiAgentMessage): message is Message {
  return message.role === "user" || message.role === "assistant" || message.role === "toolResult";
}

function getAssistantStatus(message: AssistantMessage): TaskpaneMessageStatus {
  if (message.stopReason === "aborted") {
    return "cancelled";
  }

  if (message.stopReason === "error") {
    return "error";
  }

  return "complete";
}

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
