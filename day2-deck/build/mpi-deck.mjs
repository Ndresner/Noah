import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
/* MPI and video completion by store, Jul-Sep 2026 (Noah, Oct 22 2026). Store-reported from each store's own MPI tool,
   CP + warranty ROs incl. $0; group = RO-weighted (CP+W ROs from Qlik), matching the workbook's own group line.
   One period only, so a level ranker per measure and no change ranker. MPI is scored against the 90% Roundtable
   benchmark ("90%+ MPI completion"); video and photo have no benchmark, so against the group. */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const D = JSON.parse(fs.readFileSync("./mpi/mpi.json", "utf8")).rows.map(([n, ro, mpi, mt, vid, vt, med, cpw]) => ({ n, ro, mpi, mt, vid, vt, med, cpw }));
const SEG = "Opcode Benchmarks", LBL = "JULY–SEPTEMBER 2026";
const pct = v => (v * 100).toFixed(1) + "%";
const nm = H.storeName;
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const wavg = k => { const r = D.filter(x => x[k] != null); return r.reduce((t, x) => t + x.ro * x[k], 0) / r.reduce((t, x) => t + x.ro, 0); };
const SRC = "Store-reported from each store's MPI tool (MPI_Video_Completion_Jul-Sep_2026.xlsx): CP + warranty ROs closed Jul 1 - Sep 30, 2026, including $0 ROs; % of ROs, not of MPIs. Group = RO-weighted by each store's CP+W ROs (Qlik Closed Repair Orders, 128,726). Tools differ by store; GMC Prescott's report is not filtered to CP+W; Volvo and VW St George did not say. Honda Spokane and Toyota Spokane report the same figure for MPI and video.";

function rank(k, tk, o) {
  const rows = D.filter(x => x[k] != null), G = wavg(k), ref = o.bench ?? G;
  const sorted = [...rows].sort((a, b) => b[k] - a[k]), top = sorted.filter(x => x[k] === sorted[0][k]);
  const at = rows.filter(x => x[k] >= ref - 1e-9).length, low = sorted.at(-1);
  return H.levelRanker(pptx, pptx, {
    eyebrow: SEG, title: o.title, metric: o.metric, measure: o.measure, rows: rows.map(x => [x.n, x[k]]), group: G, bench: o.bench,
    fmt: pct, groupLabel: LBL, valLabel: o.valLabel, extra: [{ label: "Tool", w: 1.3, vals: Object.fromEntries(rows.map(x => [x.n, x[tk]])) }],
    takeaway: `${at} of ${rows.length} stores at or above ${o.bench ? "the " + pct(o.bench) + " benchmark" : "the " + pct(G) + " group"}${o.bench ? "; the group is at " + pct(G) : ""}. ${list(top.map(x => nm(x.n)))} lead${top.length > 1 ? "" : "s"} at ${pct(top[0][k])}; ${nm(low.n)} is lowest at ${pct(low[k])}.`,
    source: `${SRC} ${o.def} Green at or above ${o.bench ? "the benchmark" : "the group"}, red below.${rows.length < 33 ? ` ${33 - rows.length} stores do not report it (N/A or blank).` : ""}`,
  }).addNotes(o.notes);
}

rank("mpi", "mt", { title: "MPI Completion %", metric: "mpi completion %", measure: "MPI Completion %", valLabel: "Rate", bench: 0.90,
  def: "MPI % = CP+W ROs with a completed technician multi-point inspection / CP+W ROs closed. Benchmark 90% (Roundtable: 90%+ MPI completion).",
  notes: "Every car, every time: the MPI is where the hours on the opcode slides come from. Cadillac switched to Auto.live during the quarter, so its number may understate the process." });
rank("vid", "vt", { title: "Video Completion %", metric: "video completion %", measure: "Video Completion %", valLabel: "Rate",
  def: "Video % = CP+W ROs with a technician video attached to the MPI / CP+W ROs closed. No benchmark set; scored against the group.",
  notes: `Video is what turns an inspection into a sale the customer can see. MPI at 80% or more but video under 45%, so the inspection is done and not shown: ${D.filter(x => x.mpi >= 0.8 && x.vid < 0.45).map(x => `${nm(x.n)} (${pct(x.mpi)} MPI, ${pct(x.vid)} video)`).join("; ")}.` });
rank("med", "vt", { title: "Photo Completion %", metric: "photo completion %", measure: "Photo Completion %", valLabel: "Rate",
  def: "Photo % (optional on the form) = CP+W ROs with at least one technician photo on the MPI / CP+W ROs closed. Scored against the group.",
  notes: "Five stores do not report photo completion, so the group here is weighted on the 28 stores that do (109,881 CP + warranty ROs, not the 128,726 used for MPI and video)." });

await pptx.writeFile({ fileName: "MPI & Video Completion - Review.pptx" }); console.log("mpi built", pct(wavg("mpi")), pct(wavg("vid")), pct(wavg("med")));
