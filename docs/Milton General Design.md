# AI-Native Office Runtime — Architecture Design Document

## Overview

This document summarizes the architectural decisions for a local-first, AI-native Office Add-in platform targeting:

- Excel
- Word
- PowerPoint
- potentially Outlook later

The system is intended to support:

- long-running agentic workflows
- code-native execution
- reusable macros/libraries
- local and cloud model providers
- local filesystem/model access
- collaborative document workflows
- persistent semantic memory
- browser-hosted execution
- cross-platform Office compatibility

The architecture intentionally moves away from:

- giant JSON tool registries
- server-centric orchestration
- simplistic chatbot-sidebar patterns

and toward:

- programmable document runtimes
- persistent semantic workspaces
- executable AI-native Office artifacts
- local-first agent execution

---

# Core Philosophy

The system treats Office documents as:

> executable semantic workspaces

rather than:

> static files with attached chat interfaces.

The document itself becomes:

- persistent memory substrate
- collaboration substrate
- semantic checkpoint carrier
- macro/package container
- partially self-describing computational artifact

The system follows a code-native agent architecture inspired by systems like:

- Claude Code
- OpenCode
- NanoClaw
- Pi

rather than large explicit tool-schema orchestration frameworks.

---

# High-Level Architecture

```text
Office Add-in UI
    +
Agent Runtime
    +
Model Abstraction Layer
    +
TypeScript Macro Runtime
    +
OfficeJS Adapter Layer
    +
Optional Local Capability Bridge
```

The runtime is browser-hosted inside Office WebView environments.

Execution is local-first.

Cloud infrastructure is optional.

---

# Major Technical Decisions

## 1. Code-Native Agent Execution

### Decision

The agent generates TypeScript/JavaScript programs instead of selecting among giant structured tool registries.

### Rationale

OfficeJS already provides:

- composable APIs
- batching semantics
- deferred execution
- semantic document object models

Attempting to convert OfficeJS into thousands of small structured tools would:

- explode context size
- increase latency
- reduce composability
- recreate a programming language poorly

The OfficeJS batching model strongly favors code-native execution.

### Example

```ts
export async function run(ctx) {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();

  const range = sheet.getUsedRange();
  range.load("values");

  await ctx.sync();

  // analyze values

  sheet.getRange("G1").values = [["Forecast"]];
}
```

---

# 2. TypeScript as Canonical Macro Format

### Decision

Persist macros and generated workflows as TypeScript source.

### Rationale

This provides:

- inspectability
- diffability
- replayability
- shareability
- LLM editability
- migration/versioning support
- deterministic preprocessing

The canonical artifact becomes:

```text
typed programmable source
```

rather than opaque runtime state.

### Macro Model

Macros are stored as:

```ts
export async function run(ctx) {
  ...
}
```

not serialized tool graphs.

---

# 3. Compiler/Transformation Pipeline as Security Boundary

### Decision

All generated and persisted code passes through:

```text
TS source
  -> parse
  -> static analysis
  -> AST transforms
  -> capability validation
  -> instrumentation
  -> compilation
  -> sandboxed execution
```

### Rationale

This enables:

- compile-time rejection of dangerous constructs
- validation of persisted macros
- runtime instrumentation
- deterministic execution policies
- capability enforcement

### Forbidden/Restricted Constructs

Potentially restricted:

- eval
- Function constructor
- arbitrary DOM access
- prototype mutation
- unrestricted globals
- dynamic imports
- unrestricted fetch

### Key Principle

The system executes a:

```text
safe TypeScript subset
```

rather than arbitrary browser JavaScript.

---

# 4. Controlled Runtime Surface

### Decision

Do not expose browser globals directly.

Instead expose a controlled runtime context:

```ts
export async function run(ctx: RuntimeContext) {}
```

Potential runtime surfaces:

- office
- storage
- memory
- ui
- models
- local capabilities

### Rationale

This improves:

- capability analysis
- deterministic replay
- sandboxing
- portability
- instrumentation

---

# 5. AST Instrumentation Layer

### Decision

Use AST transforms as a core runtime mechanism.

### Potential Uses

- inject tracing
- inject cancellation checks
- capability enforcement
- Office mutation tracking
- execution budgets
- replay instrumentation
- automatic batching helpers
- source rewriting

### Rationale

AST transforms provide strong runtime control while preserving:

- natural TypeScript ergonomics
- model expressiveness
- code-native workflows

---

# 6. OfficeJS as the Primary Semantic Runtime

### Decision

Treat OfficeJS itself as the domain-specific runtime.

### Rationale

OfficeJS already provides:

- semantic document models
- structured mutation APIs
- batching semantics
- transactional-ish execution

The system should leverage these strengths directly.

---

# 7. Local-First Runtime

### Decision

The system runs entirely locally by default.

Cloud infrastructure is optional.

### Rationale

Advantages:

- offline capability
- privacy
- OSS friendliness
- no required backend
- no mandatory accounts
- reduced compliance burden
- user ownership of data

---

# 8. Model Provider Architecture

### Decision

Support browser-native and local model providers.

### Provider Modes

#### Browser Direct

```text
Office Add-in
  -> provider API
```

Examples:

- OpenAI
- Anthropic
- OpenRouter
- Groq
- Together

using BYOK semantics.

#### Local Companion

```text
Office Add-in
  -> localhost bridge
  -> local models
```

Examples:

- Ollama
- LM Studio
- vLLM

### Rationale

This preserves:

- provider flexibility
- local/private inference
- OSS deployment simplicity
- model portability

---

# 9. Reuse of Pi Agent Architecture

### Decision

Reuse/adapt substantial portions of the Pi agent architecture.

### Expected Reuse Areas

- agent loop
- model abstraction layer
- streaming orchestration
- memory/checkpointing
- transcript handling
- async task model

### Expected Rewrite Areas

- runtime substrate
- Office-specific execution
- filesystem assumptions
- browser/runtime assumptions
- UI/session integration

### Architectural Direction

The architecture becomes:

```text
Pi agent core
    +
TS macro runtime
    +
OfficeJS adapter
    +
local capability bridge
```

---

# 10. Soft Fork / Workspace Strategy

### Decision

Fork Pi as a distinct package/workspace rather than hard vendoring.

### Proposed Structure

```text
packages/
  pi-core/
  runtime-office/
  runtime-browser/
  runtime-local/
  ui-office/
```

### Rationale

This preserves:

- architectural boundaries
- upstream mergeability
- subsystem separation
- cleaner dependency structure

Goal:

```text
diverge cleanly
```

rather than:

```text
copy random source into app/
```

---

# 11. Browser-Compatible Model Abstraction Layer

### Decision

Adapt the model abstraction layer to be browser-first.

### Key Principle

Avoid Node.js runtime assumptions inside the core layer.

### Requirements

- Web Streams support
- browser-compatible async iteration
- transport abstraction
- provider abstraction
- streaming-first design
- cancellation support

### Rationale

The Office runtime is browser-hosted.

The core should remain:

- environment-agnostic
- transport-agnostic
- provider-agnostic

---

# 12. Local Capability Bridge

### Decision

Expose advanced capabilities through an optional local companion process.

### Architecture

```text
Office Add-in
    ↕
localhost bridge
    ↕
filesystem/models/OS services
```

### Initial Transport

Simple typed RPC over:

- localhost HTTP
- WebSocket

### Why Not MCP Initially

MCP is unnecessary for initial architecture.

A typed capability API is sufficient.

MCP compatibility may be added later if:

- ecosystem interoperability
- third-party extensions
- pluggable capability providers

become strategically important.

---

# 13. Capability-Oriented Local APIs

### Decision

Expose controlled capability APIs rather than raw shell access.

### Example

Preferred:

```ts
workspace.readFile()
workspace.search()
workspace.writeFile()
```

Avoid:

```text
raw unrestricted shell execution
```

### Rationale

This improves:

- security
- auditability
- portability
- determinism
- replayability

---

# 14. Persistent Memory Architecture

### Decision

Separate:

- conversation history
- execution/event logs
- semantic checkpoints
- retrieval caches

### Memory Layers

| Layer | Purpose |
|---|---|
| Conversation transcript | Durable user-facing history |
| Execution log | Runtime/tool traces |
| Semantic checkpoints | Compact structured memory |
| Retrieval cache | Large local retrieval/indexing |

### Rationale

Users expect:

- prompts/history to persist

but do not necessarily care about preserving:

- low-level tool traces forever.

---

# 15. Event-Sourced Agent Memory

### Decision

Use append-oriented execution/event logs.

### Example

```json
{"type":"user","content":"Analyze workbook"}
{"type":"tool_call","tool":"read_range"}
{"type":"tool_result","result":...}
```

### Rationale

Advantages:

- replayability
- debugging
- resumability
- deterministic reconstruction
- semantic checkpointing

---

# 16. Structured Semantic Checkpoints

### Decision

Compact memory into structured semantic state.

Avoid relying primarily on prose summaries.

### Example

```json
{
  "document_model": {...},
  "active_tasks": [...],
  "known_entities": [...],
  "workflow_state": {...}
}
```

### Rationale

Structured checkpoints are:

- more stable
- more compressible
- more replayable
- more migratable
- less lossy

than freeform summaries.

---

# 17. Document-Embedded Shared State

### Decision

Persist shared semantic state inside Office documents.

### Likely Mechanism

Custom XML Parts.

### Stored Data

- semantic checkpoints
- workflow state
- compact memory
- recent context
- transcript metadata

### Not Stored

Avoid storing:

- API keys
- large embeddings
- giant raw execution traces

### Rationale

The document becomes:

```text
shared canonical collaborative state
```

across co-authoring sessions.

---

# 18. Local vs Shared State Separation

### Shared State

Stored in document:

- transcript
- semantic checkpoints
- workflow memory
- collaborative state

### Local State

Stored locally:

- embeddings
- retrieval indexes
- execution caches
- filesystem state
- local model caches

### Rationale

This aligns naturally with Office co-authoring semantics.

---

# 19. IndexedDB as Local Runtime Cache

### Decision

Use IndexedDB for:

- execution caches
- embeddings
- retrieval indexes
- compiled artifacts
- local semantic memory

### Rationale

IndexedDB provides:

- large storage capacity
- browser compatibility
- offline persistence
- efficient querying

---

# 20. Compilation Artifact Caching

### Decision

Cache compiled/instrumented artifacts.

### Model

```text
TS source (canonical)
    ↓
compiled/instrumented JS (cache)
    ↓
sandbox execution
```

### Cache Strategy

Use content-addressed caching:

```text
hash(
  source
  + runtime_version
  + stdlib_version
  + transform_pipeline_version
)
```

### Rationale

Avoid repeated:

- parsing
- instrumentation
- compilation
- dependency analysis

---

# 21. Ribbon Strategy

### Decision

Treat the ribbon as a lightweight launcher surface.

### Use Ribbon For

- opening taskpane
- launching workflows
- contextual entry points
- pinned actions

### Use Taskpane For

- dynamic workflows
- macro management
- execution UI
- agent interaction
- semantic exploration

### Runtime Ribbon Creation

Use:

```ts
Office.ribbon.requestCreateControls(...)
```

for contextual runtime-created tabs where useful.

### Rationale

Taskpanes are significantly better suited for:

- dynamic agent systems
- rich interactions
- evolving workflows

than ribbon UIs.

---

# 22. Co-Authoring Model

### Key Principle

Each collaborator runs an independent runtime.

There is no globally shared live agent process.

### Shared Layer

The Office document itself becomes the synchronization substrate.

### Implications

- append-oriented logs preferred
- semantic checkpoints preferred
- operation-oriented synchronization preferred
- local execution remains per-user

### Rationale

This aligns with how Office co-authoring actually functions.

---

# 23. Execution Philosophy

The system intentionally favors:

```text
programmable runtime
```

over:

```text
large explicit tool registries
```

and:

```text
semantic object models
```

over:

```text
vision-driven computer use
```

wherever possible.

OfficeJS is treated as the primary semantic automation substrate.

GUI automation is considered a fallback mechanism.

---

# Long-Term Direction

The architecture is ultimately converging toward:

```text
AI-native programmable Office runtime
```

featuring:

- persistent semantic documents
- executable macros/workflows
- local-first agent execution
- typed programmable automation
- reusable capability libraries
- collaborative semantic memory
- browser-native execution
- optional local compute augmentation

This is conceptually closer to:

- notebook systems
- programmable IDE runtimes
- AI operating environments

than traditional Office add-ins.

