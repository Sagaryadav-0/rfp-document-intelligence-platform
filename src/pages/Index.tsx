import { supabase } from "@/integrations/supabase/client";
import { useCallback, useRef, useState } from "react";
import { FileText, UploadCloud, Loader2, Download, CheckCircle2, AlertCircle, Sparkles } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import { getSessionId } from "@/lib/session";
import * as XLSX from "xlsx";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";

type Status = "idle" | "ready" | "processing" | "done" | "error";

const Index = () => {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [result, setResult] = useState<{ url: string; name: string; sections: number; pages: number } | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [pageRanges, setPageRanges] = useState<string>("");
  const inputRef = useRef<HTMLInputElement>(null);

  type PageRange = { from: number; to: number };

  const parsePageRanges = (input: string | null | undefined): PageRange[] => {
    if (!input) return [];
    const normalized = input
      .replace(/[;]+/g, ",")
      .replace(/\s*,\s*/g, ",")
      .trim();
    const out: PageRange[] = [];
    for (const part of normalized.split(",")) {
      const trimmed = part.trim();
      let m = trimmed.match(/^(\d+)\s*(?:-|–|—|to)\s*(\d+)$/i);
      if (!m) m = trimmed.match(/^(\d+)$/);
      if (!m) continue;
      const from = parseInt(m[1], 10);
      const to = m[2] ? parseInt(m[2], 10) : from;
      if (from > 0 && to >= from) out.push({ from, to });
    }
    return out;
  };

  const pageInRanges = (page: number, ranges: PageRange[]) => {
    if (ranges.length === 0) return true;
    return ranges.some((r) => page >= r.from && page <= r.to);
  };

  const acceptFile = useCallback((f: File | null | undefined) => {
    if (!f) return;
    if (!/\.pdf$/i.test(f.name)) {
      toast.error("Please upload a PDF file.");
      return;
    }
    if (f.size > 25 * 1024 * 1024) {
      toast.error("PDF must be under 25 MB.");
      return;
    }
    setFile(f);
    setStatus("ready");
    setErrorMsg("");
    setResult(null);
  }, []);

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    acceptFile(e.dataTransfer.files?.[0]);
  };

  const buildLocalWorkbook = (rows: string[][]) => {
    const sheetData: (string | number)[][] = [["S.No.", "Technical Specifications", "Compliance (Yes/No)", "Remarks"]];
    rows.forEach((row, index) => {
      sheetData.push([index + 1, row[0], "", ""]);
    });

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws["!cols"] = [
      { wch: 8 },
      { wch: 70 },
      { wch: 22 },
      { wch: 36 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Compliance");
    return XLSX.write(wb, { bookType: "xlsx", type: "array" });
  };

  const extractParagraphs = (text: string) => {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  const paragraphs: string[] = [];
  let current = "";

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (!current) {
      current = line;
      continue;
    }

    // If previous line ends with sentence punctuation → new paragraph
    if (/[.!?:]$/.test(current)) {
      paragraphs.push(current.trim());
      current = line;
    } else {
      // same paragraph (wrapped line)
      current += " " + line;
    }
  }

  if (current && current.length > 30) {
    paragraphs.push(current.trim());
  }

  return paragraphs;
};

  const isPdfHeaderFooterLine = (line: string) => {
    const trimmed = line.trim();
    if (!trimmed) return false;
    const patterns = [
      /^page\s*\d+(?:\s*of\s*\d+)?$/i,
      /\bconfidential\b/i,
      /\bministry of information\b/i,
      /\bscope of work\b/i,
      /\bDTM-T-[A-Z0-9-]+\b/i,
    ];
    return patterns.some((re) => re.test(trimmed));
  };

  const localPdfFallback = async () => {
    const pdfjsLib = await import("pdfjs-dist/legacy/build/pdf");
    pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    const buffer = await file.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: buffer });
    const pdf = await loadingTask.promise;
    const pageCount = pdf.numPages;
    const ranges = parsePageRanges(pageRanges);
    const selectedPages = Array.from({ length: pageCount }, (_, i) => i + 1).filter((page) => pageInRanges(page, ranges));

    const rows: string[][] = [];
    for (const pageNum of selectedPages) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      let text = (content.items as Array<{ str: string }>).
        map((item) => item.str)
        .join("\n");
      const pageLines = text
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean)
        .filter((line) => !isPdfHeaderFooterLine(line));
      text = pageLines.join(" ")
        .replace(/(?:DTM-T-[A-Z0-9-]+|Ministry of Information Digital Platform Confidential|C4 Scope of Work)/gi, " ")
        .replace(/Page\s*\d+(?:\s*of\s*\d+)?/gi, " ")
        .replace(/\s+/g, " ")
        .trim();
      // Normalize spacing
      // Add logical paragraph breaks
      text = text.replace(/([.!?:])\s+(?=[A-Z0-9])/g, "$1\n");
      const paragraphs = extractParagraphs(text);
      paragraphs.forEach((para) => {
        rows.push([para]);
      });
    }

    const workbookArray = buildLocalWorkbook(rows);
    const blob = new Blob([workbookArray], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return {
      url: URL.createObjectURL(blob),
      name: file.name.replace(/\.pdf$/i, "") + " - local-extract.xlsx",
      sections: rows.length,
      pages: pageCount,
    };
  };

  const uploadToStorage = async (file: File) => {
  const filePath = `uploads/${Date.now()}_${file.name}`;

  const { error } = await supabase.storage
    .from("rfp-files")
    .upload(filePath, file);

  if (error) {
    throw new Error(error.message);
  }

  return filePath;
};

  const process = async () => {
    if (!file) return;
    setStatus("processing");
    setErrorMsg("");

    const tryLocalFallback = async (reason: string) => {
      try {
        const localResult = await localPdfFallback();
        setResult(localResult);
        setStatus("done");
        toast.success(`Excel ready locally after remote fallback: ${reason}`);
      } catch (localErr) {
        const localMsg = localErr instanceof Error ? localErr.message : "Local fallback failed.";
        setErrorMsg(`${reason} — ${localMsg}`);
        setStatus("error");
        toast.error(`Failed to process locally: ${localMsg}`);
      }
    };

    try {
      const sessionId = getSessionId();

      const formData = new FormData();
      formData.append("file", file);
      formData.append("session_id", sessionId);
      formData.append("pageRanges", pageRanges);

      const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string;
      const appToken = typeof window !== "undefined" ? sessionStorage.getItem("rfp-app-token") : null;

      const headers: Record<string, string> = {};
      if (appToken) {
        headers["x-app-token"] = appToken;
      } else {
        const {
          data: { session },
        } = await supabase.auth.getSession();

        const accessToken = session?.access_token;

        if (!accessToken) {
          throw new Error("User not authenticated");
        }

        headers["Authorization"] = `Bearer ${accessToken}`;
      }

      const res = await fetch(`${supabaseUrl}/functions/v1/process-rfp`, {
        method: "POST",
        body: formData,
        headers,
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || "Function failed");
      }

      const { excelUrl, excelBase64, fileName, sections, pages } = data;

      let downloadUrl = excelUrl ?? null;

      if (!downloadUrl && excelBase64) {
        const bytes = Uint8Array.from(atob(excelBase64), (c) => c.charCodeAt(0));
        const blob = new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        downloadUrl = URL.createObjectURL(blob);
      }

      if (!downloadUrl) {
        throw new Error("No Excel file was returned from the server.");
      }

      setResult({ url: downloadUrl, name: fileName, sections, pages });
      setStatus("done");
      toast.success("Excel ready to download.");
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Something went wrong.";
      if (msg.includes("Failed to fetch") || msg.includes("Remote function not found")) {
        await tryLocalFallback(msg);
        return;
      }
      setErrorMsg(msg);
      setStatus("error");
      toast.error(msg);
    }
  };

  const reset = () => {
    if (result?.url?.startsWith("blob:")) {
      URL.revokeObjectURL(result.url);
    }
    setFile(null);
    setStatus("idle");
    setErrorMsg("");
    setPageRanges("");
    setResult(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <main className="min-h-screen bg-[image:var(--gradient-soft)]">
      <section className="container mx-auto px-4 py-12 md:py-20">
        <header className="mx-auto max-w-3xl text-center">
          <div className="mx-auto mb-6 flex items-center gap-3 rounded-full border border-border bg-card px-4 py-2 text-sm text-muted-foreground shadow-sm">
            <img src="/logo.svg" alt="Logo" className="h-8 w-8" />
            <span className="font-semibold">Compliance Project</span>
          </div>
          <h1 className="mb-4 text-balance text-4xl font-bold tracking-tight text-foreground md:text-5xl">
            RFP to Compliance Excel
          </h1>
          <p className="text-balance text-lg text-muted-foreground">
            Upload your RFP and get a structured Excel with technical specifications extracted as individual compliance items, 
            ready for manual review and approval.
          </p>
        </header>

        <Card
          className="mx-auto mt-10 max-w-3xl overflow-hidden border-border/60 shadow-[var(--shadow-card)]"
        >
          <div className="bg-[image:var(--gradient-hero)] px-8 py-6 text-primary-foreground">
            <div className="flex items-center gap-3">
              <FileText className="h-6 w-6" />
              <div>
                <h2 className="text-lg font-semibold">RFP Compliance Extractor</h2>
                <p className="text-sm opacity-90">PDF in. Structured compliance Excel out.</p>
              </div>
            </div>
          </div>

          <div className="p-6 md:p-8">
            {status !== "done" && (
              <div
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={onDrop}
                onClick={() => inputRef.current?.click()}
                className={`flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed px-6 py-12 text-center transition-colors ${
                  dragOver ? "border-primary bg-accent" : "border-border hover:border-primary/50 hover:bg-accent/50"
                }`}
              >
                <UploadCloud className="mb-3 h-10 w-10 text-primary" />
                <p className="text-base font-medium text-foreground">
                  {file ? file.name : "Drop your RFP PDF here"}
                </p>
                <p className="mt-1 text-sm text-muted-foreground">
                  {file ? `${(file.size / 1024 / 1024).toFixed(2)} MB` : "or click to browse — PDF up to 25 MB"}
                </p>
                <input
                  ref={inputRef}
                  type="file"
                  accept="application/pdf,.pdf"
                  className="hidden"
                  onChange={(e) => acceptFile(e.target.files?.[0])}
                />
              </div>
            )}

            {status === "ready" && (
              <div className="mt-6 space-y-4">
                <div>
                  <label htmlFor="page-ranges" className="mb-1.5 block text-sm font-medium text-foreground">
                    Page ranges for compliance <span className="font-normal text-muted-foreground">(optional)</span>
                  </label>
                  <input
                    id="page-ranges"
                    type="text"
                    value={pageRanges}
                    onChange={(e) => setPageRanges(e.target.value)}
                    placeholder="e.g. 5-12, 20-35, 50"
                    className="w-full rounded-md border border-border bg-background px-3 py-2 text-sm shadow-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
                  />
                  <p className="mt-1 text-xs text-muted-foreground">
                    Leave blank to include the whole PDF. Use commas to add multiple ranges.
                  </p>
                </div>
                <div className="flex flex-col-reverse gap-3 sm:flex-row sm:justify-end">
                  <Button variant="ghost" onClick={reset}>Choose another file</Button>
                  <Button onClick={process} size="lg" className="shadow-[var(--shadow-elegant)]">
                    <Sparkles className="mr-2 h-4 w-4" />
                    Extract Compliance Items
                  </Button>
                </div>
              </div>
            )}

            {status === "processing" && (
              <div className="mt-8 flex flex-col items-center justify-center gap-3 py-6 text-center">
                <Loader2 className="h-8 w-8 animate-spin text-primary" />
                <p className="font-medium text-foreground">Extracting technical specifications from RFP…</p>
                <p className="text-sm text-muted-foreground">This usually takes 10–30 seconds.</p>
              </div>
            )}

            {status === "error" && (
              <div className="mt-6 rounded-lg border border-destructive/40 bg-destructive/10 p-4">
                <div className="flex items-start gap-3">
                  <AlertCircle className="mt-0.5 h-5 w-5 text-destructive" />
                  <div className="flex-1">
                    <p className="font-medium text-foreground">Couldn't process this PDF</p>
                    <p className="mt-1 text-sm text-muted-foreground">{errorMsg}</p>
                  </div>
                  <Button variant="outline" size="sm" onClick={reset}>Try again</Button>
                </div>
              </div>
            )}

            {status === "done" && result && (
              <div className="flex flex-col items-center justify-center gap-4 py-6 text-center">
                <div className="flex h-14 w-14 items-center justify-center rounded-full bg-primary/10">
                  <CheckCircle2 className="h-7 w-7 text-primary" />
                </div>
                <div>
                  <p className="text-lg font-semibold text-foreground">Compliance Excel Ready</p>
                  <p className="text-sm text-muted-foreground">
                    Generated <span className="font-medium text-foreground">{result.sections}</span> technical specifications from the RFP, 
                    formatted as individual compliance items ready for manual review.
                  </p>
                </div>
                <div className="flex flex-col gap-2 sm:flex-row">
                  <Button asChild size="lg" className="shadow-[var(--shadow-elegant)]">
                    <a href={result.url} download={result.name}>
                      <Download className="mr-2 h-4 w-4" />
                      Download Compliance Excel
                    </a>
                  </Button>
                  <Button variant="outline" onClick={reset}>Process Another RFP</Button>
                </div>
              </div>
            )}
          </div>
        </Card>

        <section className="mx-auto mt-12 grid max-w-3xl gap-4 sm:grid-cols-3">
          {[
            { title: "Technical Specifications", body: "Each paragraph from the RFP becomes a separate compliance item with S.No." },
            { title: "Manual Review Ready", body: "Compliance (Yes/No) and Remarks columns left blank for manual evaluation." },
            { title: "Structured Format", body: "Clean Excel format with proper column widths and filename-based sheet naming." },
          ].map((f) => (
            <Card key={f.title} className="border-border/60 p-5 shadow-[var(--shadow-card)]">
              <h3 className="font-semibold text-foreground">{f.title}</h3>
              <p className="mt-1 text-sm text-muted-foreground">{f.body}</p>
            </Card>
          ))}
        </section>
      </section>
    </main>
  );
};
export default Index;