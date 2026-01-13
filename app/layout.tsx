import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Financial Model Generator",
  description: "Generate downloadable 3-statement financial models from live market data."
};

export default function RootLayout({
  children
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-slate-950 text-slate-100">
        {children}
      </body>
    </html>
  );
}
