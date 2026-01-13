"use client";

import { useState } from "react";

export default function HomePage() {
  const [ticker, setTicker] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError(null);

    const trimmedTicker = ticker.trim().toUpperCase();
    if (!trimmedTicker) {
      setError("Please enter a stock ticker.");
      return;
    }

    setIsLoading(true);
    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ticker: trimmedTicker })
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload?.error || "Unable to generate model.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${trimmedTicker}-financial-model.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      const message = err instanceof Error ? err.message : "Something went wrong.";
      setError(message);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <main className="flex min-h-screen items-center justify-center px-6 py-16">
      <div className="w-full max-w-2xl space-y-10 rounded-3xl bg-slate-900/70 p-10 shadow-xl ring-1 ring-slate-800">
        <header className="space-y-4 text-center">
          <p className="text-sm font-semibold uppercase tracking-[0.3em] text-slate-400">
            Financial Model Generator
          </p>
          <h1 className="text-4xl font-semibold text-white">
            Build downloadable 3-statement models in seconds
          </h1>
          <p className="text-base text-slate-300">
            Enter any public ticker and we will assemble a polished Excel workbook with financial
            statements, valuation multiples, and driver assumptions.
          </p>
        </header>

        <form onSubmit={handleSubmit} className="space-y-6">
          <div className="flex flex-col gap-4 sm:flex-row">
            <input
              type="text"
              value={ticker}
              onChange={(event) => setTicker(event.target.value)}
              placeholder="e.g. AAPL"
              className="flex-1 rounded-2xl border border-slate-700 bg-slate-950 px-4 py-3 text-lg text-white placeholder:text-slate-500 focus:border-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-400/40"
              aria-label="Stock ticker"
            />
            <button
              type="submit"
              disabled={isLoading}
              className="rounded-2xl bg-blue-500 px-6 py-3 text-base font-semibold text-white transition hover:bg-blue-400 disabled:cursor-not-allowed disabled:bg-blue-500/60"
            >
              {isLoading ? "Generating..." : "Generate Model"}
            </button>
          </div>
          {error ? (
            <p className="rounded-xl border border-red-400/40 bg-red-500/10 px-4 py-3 text-sm text-red-200">
              {error}
            </p>
          ) : null}
        </form>

        <section className="grid gap-4 text-sm text-slate-300 sm:grid-cols-2">
          <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
            <h2 className="text-base font-semibold text-white">Single API call</h2>
            <p>We fetch everything from Yahoo Finance with one quoteSummary request.</p>
          </div>
          <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
            <h2 className="text-base font-semibold text-white">Excel-ready output</h2>
            <p>Statements, ratios, and assumptions are formatted and ready to edit.</p>
          </div>
        </section>
      </div>
    </main>
  );
}
