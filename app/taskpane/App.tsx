import { useState } from "react";
import { insertText } from "./excel";

export default function App() {
  const [text, setText] = useState("Ask Milton to analyze this workbook.");
  const [status, setStatus] = useState<string | null>(null);

  async function handleInsert() {
    setStatus("Writing to the active worksheet...");

    try {
      await insertText(text);
      setStatus("Inserted into cell A1.");
    } catch (error) {
      console.error(error);
      setStatus("Excel rejected the write. Check the dev console for details.");
    }
  }

  return (
    <main className="shell">
      <header className="hero">
        <img className="logo" src="/assets/logo-filled.png" alt="" />
        <div>
          <p className="eyebrow">Excel add-in</p>
          <h1>Milton</h1>
          <p className="lede">Bring-your-own-model spreadsheet workflows, starting with a small Office API smoke test.</p>
        </div>
      </header>

      <section className="panel" aria-labelledby="insert-heading">
        <h2 id="insert-heading">Write to the workbook</h2>
        <label htmlFor="text-to-insert">Text</label>
        <textarea id="text-to-insert" value={text} onChange={(event) => setText(event.target.value)} />
        <button type="button" onClick={handleInsert}>
          Insert into A1
        </button>
        {status ? <p className="status">{status}</p> : null}
      </section>
    </main>
  );
}
