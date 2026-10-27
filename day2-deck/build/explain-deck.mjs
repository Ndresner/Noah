/* Metric explainer slides for review (2026-10-19). Not in the deck yet. Definitions match how each number on the
   data slides is calculated. Worked examples: Service Manager Roundtable (Nov 2024) slide 21 table, and simple
   illustrations. Group numbers: the verified Jan-Sep 2026 figures on the review deck and proficiency slides. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SRC = "Definitions match the data slides: Qlik - Closed Repair Orders (all customer-pay ROs, by close month), Profit & Loss and Productivity apps, Jan-Sep 2026 vs 2025, same-store, 33 stores. Worked example: Service Manager Roundtable, Nov 15, 2024, slide 21.";
const X = (o) => H.explainer(pptx, pptx, { source: SRC, ...o });
const RT = "Roundtable example";

X({ eyebrow: "Customer Pay", title: "CP RO Count",
  formula: ["CP RO count", "=", "customer-pay repair orders closed in the period"], formulaSize: 22,
  define: "Every RO with a customer-pay line. One RO counts once, however many lines it has.",
  why: ["Everything else starts here: hours, ELR and $ per RO all multiply off the RO count.",
        "It shows whether the business is expanding or contracting: are customers coming back?",
        "Flat car count means growth has to come from selling more on each visit."],
  levers: [["Appointment availability", "Same-day and next-day slots; an easy online scheduler"],
           ["Phone handling", "Every call answered, and every caller asked for the appointment"],
           ["Retention", "Proactive outreach and ongoing engagement: lapsed, inactive and sold-not-serviced customers (VizAuto retention app)"],
           ["Next visit booked", "First appointment at the sale; next one before the customer leaves"],
           ["Quick maintenance", "Maintenance available on demand, done quickly and efficiently, keeps customers coming back"],   // no "express" in the deck (Noah, Oct 26)
           ["Marketing", "Fixed ops budget, keyword search, OEM and third-party reminders"]],
  leverRows: [1, 1.35, 1],   // the retention line runs four lines (Oct 23)
  ours: "Group: 279,756 CP ROs, up 456 from 2025. Volume is flat.",
  notes: "PLACE: CP Engine, in front of CP RO Count.\nDefinition is the data slides' own: distinct ROs with a customer-pay line, by close month. This includes the roughly 0.9% of CP ROs with no labor charged (parts-only sales such as additives, wipers and batteries, and blank CP lines); Qlik's hasLabor flag is set on every line, so it does not filter them, and the count is kept that way to match the numbers stores are used to. GMC Prescott body-shop-only ROs excluded." });

X({ eyebrow: "Customer Pay", title: "Hours per RO",
  formula: ["Hours per RO", "=", "CP hours sold", "÷", "CP RO count"],
  define: "The labor hours we sell, on average, every time a customer-pay car comes through the drive.",
  example: [["0.1 hour more on every RO", "= 27,976 more hours"], ["× $162 an hour (our ELR)", "= $4.5M more labor"]], exampleLabel: "0.1 HOUR",
  // Oct 26 (Noah): "the work we sell"; it shows where a lift came from (price or more work)
  why: ["It cannot be manipulated: it is the work we sell.",
        "It shows where the lift came from: if labor sales per RO go up but hours per RO stay flat, we are charging more, not selling more.",
        "Selling more hours also sells parts."],
  levers: [["MPI on every car", "Every car inspected, with photos and video sent to the customer"],
           ["Road to the Sale", "Walk-around at the vehicle: hood, tires and wipers"],
           ["Menu presentation", "Time and mileage recommendations, every customer, every time"],
           ["Declined work follow-up", "A defined process and a champion to bring it back"],
           ["Op code times", "Sell the hours the job pays; don't give time away"],
           ["Right tech, right car", "High-mileage and repair work to the techs who find it"]],
  ours: "Group: 1.87 CP hours per RO, up 0.07 and ahead of 2025 in all nine months.",
  tone: "g", notes: "PLACE: CP Engine, in front of CP Hours per RO.\n$4,543,691 = 0.1 x 279,756 CP ROs (27,976 hours) x $162.42 CP ELR (on the Hours per RO data slide). Slide rounds to $162 and $4.5M so the room can follow it (Noah, Oct 23)." });

X({ eyebrow: "Customer Pay", title: "Effective Labor Rate",
  formula: ["Effective labor rate", "=", "CP labor sale", "÷", "CP hours sold"],
  define: "The average amount we collect for each hour of customer-pay labor we sell.",
  // Oct 26 (Noah): verify changes; look at it by work type and op code, not just the total
  why: ["It is what the shop actually collects per hour, and it shows whether a pricing change did what it was meant to do.",
        "Look past the total: by work type and by op code there can be room to raise a rate, or a reason to lower one.",
        "Benchmark: CP ELR at least 70% of door rate."],
  // Oct 23 (Noah, from the one-slide preview): work-mix points live in What Moves It; door rate and labor grid combined
  levers: [["Work mix", "Repair is priced above maintenance: repair-heavy shops run a higher ELR, maintenance-heavy lower. Neither is wrong."],
           ["Maintenance pricing", "If most of your hours are maintenance, the maintenance rate moves ELR more than the repair rate"],
           ["ELR by labor operation", "Know what every op code collects per hour, not just the store total"],
           ["Door rate and repair pricing", "Posted rate from a market labor study, and how repair jobs are priced against it"],
           ["Discounts", "Every discount lowers ELR. Run them through ADC so we can see them."],
           ["Maintenance labor times", "The hours we set and pay the tech on menu jobs; repair times come from labor guides"]],
  leverRows: [1.25, 1, 1],
  ours: "Group: $162.42 CP ELR, up $3.56. Jul-Sep it ran 74.7% of door rate; 25 of 33 stores at or above 70%.",
  tone: "g", notes: "PLACE: Customer Pay, after CP Hours per RO and ahead of Labor Sale by Pay Type and the CP ELR slides.\nRoundtable worked example removed from the slide (Noah, Oct 21).\nTalk track (Noah, Oct 23): to understand your ELR you have to understand your work mix. A shop that does more repair than maintenance will run a higher ELR, because the repair rate is higher; a shop that is mostly maintenance will run lower. There is nothing wrong with that, but you need to know your mix to know why your ELR is what it is. Raising the repair rate may do less than raising the maintenance rate if most of your work is maintenance. Know your ELR for every labor operation." });

X({ eyebrow: "Customer Pay", title: "CP Labor $ per RO",
  formula: ["CP labor $ per RO", "=", "hours per RO", "×", "ELR"],
  define: "Customer-pay labor sale divided by CP RO count. It combines how much we sell with what we collect for it.",
  example: [["1.87 hours per RO", "× $162.42"], ["= CP labor per RO", "$304"]], exampleLabel: "OUR NUMBERS",
  why: ["Focus on CP labor dollars per RO and every other metric rises with it.",   // Noah, Oct 25
        "With additional labor sales come additional parts sales.",
        "It is the single best metric to measure a service advisor's performance."],
  // Oct 26 (Noah): specific levers, no hours-per-RO cross reference; pay plans that reward what advisors control
  levers: [["Written process", "Follow it on every car: a multi-point inspection every time"],
           ["Service menu", "Present the menu and its recommendations, consistently, every customer"],
           ["Pay plans", "Reward advisors for CP labor sales and gross: what they control, and a true measure of their performance"],
           ["Tracking and accountability", "What gets tracked gets improved: accountability through visibility"]],   // replaces the ELR lever (Noah, Oct 23)
  ours: "Group: $304 CP labor per RO, up $17. With ROs flat, the $4,899,376 labor gain came from selling more per visit.",
  tone: "g", notes: "PLACE: CP Engine, in front of CP Labor $ per RO.\n1.87 x $162.42 = $303.73, which rounds to the $304 on the data slide. Parts line (Noah, Oct 23: the parts $ per RO section is out of the deck, so the link lives here): CP parts per RO $228.50 -> $237.96 Jan-Sep 2025 -> 2026 while CP labor per RO rose $286.67 -> $303.72 (Qlik Closed RO, m5.json)." });

X({ eyebrow: "Customer Pay", title: "CP Parts $ per RO",
  formula: ["CP parts $ per RO", "=", "CP parts sale", "÷", "CP RO count"],
  define: "Customer-pay parts sold through the service drive, per visit.",
  why: ["Parts follow labor: sell the hour and the parts come with it.",
        "It adds parts gross on the same car count.",
        "Luxury stores run more than twice the rest, so compare a store with its own segment."],
  levers: [["Hours per RO", "More repair and maintenance hours carry more parts"],
           ["Menu bundles", "Group-priced maintenance parts for dealer and customer"],
           ["Parts pricing matrix", "What we charge relative to cost"],
           ["Staging and special orders", "Parts ready at the bay; special orders followed up"]],
  ours: "Group: $238 CP parts per RO, up $9 and ahead of 2025 in all nine months.",
  tone: "g", notes: "PLACE: CP Engine, in front of CP Parts $ per RO." });

X({ eyebrow: "Customer Pay", title: "Labor Gross Profit %",
  formula: ["Labor GP %", "=", "(CP labor sale", "−", "tech cost)", "÷", "CP labor sale"],
  define: "The share of each labor dollar we keep after paying the technician.",
  example: [["Sell 1 hour at", "$162"], ["Pay the tech", "$31"], ["We keep", "$131"], ["$131 ÷ $162", "81%"]], exampleLabel: "ONE HOUR SOLD",
  why: ["The target range is 78% to 80%: not too high, not too low. Hold it by matching the tech to the work, not by squeezing tech pay.",   // back to 78-80 (Noah, Oct 26)
        "Too low means our labor price is too low for what we pay the tech, or tech cost is too high for what we charge.",
        "Too high can mean techs are underpaid, which can cost us technicians and hurt retention."],
  levers: [["Effective labor rate", "Same tech cost, higher ELR = higher margin"],
           ["Tech pay rate", "Flat-rate pay per hour against the rate we sell"],
           ["Right tech, right work", "Skill-code costing: don't pay a master-tech rate on an oil change"],
           ["Type of work", "Maintenance runs a lower margin than repair at full door rate"],
           ["Low-rate work", "Internal detail and other low-dollar-per-hour work pull the margin down"],
           ["Discounting", "Monitor it and keep controls in place"]],   // comebacks out (Noah, Oct 26)
  ours: "Group: 80.8% CP labor gross profit, down 0.1 pts and just above the 78-80% target range.",
  notes: "PLACE: CP Engine, in front of CP Labor Gross Profit %.\nExample (Noah, Oct 23: replaces the Roundtable table): one CP hour at the group's Jan-Sep 2026 averages, $162.42 ELR, $31.21 tech cost per hour sold, $131.21 kept = 80.8%; rounded on the slide.\nOct 23 (Noah): The Goldilocks Margin talking point is folded into this slide (it repeated it): no Goldilocks or Roundtable wording; its ask (match the tech to the job, not squeeze tech pay) is the first Why It Matters point." });

X({ eyebrow: "Customer Pay", title: "Parts Gross Profit %",
  formula: ["Parts GP %", "=", "(CP parts sale", "−", "parts cost)", "÷", "CP parts sale"],
  define: "The share of each customer-pay parts dollar we keep, less the cost of the part.",
  // Oct 26 (Noah): the pricing matrix is the biggest tool; full talk track in the notes
  // Oct 26 (Noah): reasons from his "5 reasons parts GP % matters"; the full list is in the notes
  why: ["Parts keeps about 37¢ of every gross dollar after its own expenses, the most of any department (service 22¢): margin goes almost straight to the bottom line.",   // Noah, Oct 26
        "It shows whether parts sell at the right price: matrix, markups and discounts.",
        "It flags heavy discounting, or cost increases we have not passed on.",
        "CP parts pricing supports the warranty parts reimbursement rate."],
  levers: [["Parts pricing matrix", "Know how it is set up, make sure pricing follows the matrix setup, and review it against your gross goals"],
           ["Discounts and overrides", "Monitor discounts and manual price changes; keep safeguards in place"],
           ["Stay competitive", "Balance margin with price, especially on high-cost parts"],
           ["Mix", "Tires and competitive items carry thinner margins"],
           ["Menu pricing", "Bundled maintenance parts priced to protect margin"],
           ["Tracking", "Monitor and actively manage your parts GP %: know where you stand and adjust as needed"]],   // Noah, Oct 26
  leverRows: [1.35, 1, 1],
  ours: "Group: 40.7% CP parts gross profit, up 1.2 pts and ahead of 2025 in all nine months.",
  tone: "g", notes: "PLACE: CP Engine, in front of CP Parts Gross Profit %.\nWhy understand your parts pricing matrix (Noah, Oct 26): (1) Maximize gross profit: price parts to maximize profitability; find gross without more volume. (2) Consistent pricing: every customer charged the same for the same part; no incorrect prices from bad setups or manual overrides. (3) Stay competitive: avoid markups that make repairs too expensive; balance profit with price on higher-cost parts. (4) Check the matrix is applied: parts actually sell at matrix prices; find parts, labor types or transactions that bypass it. (5) Warranty reimbursement: customer-pay parts pricing drives the warranty parts reimbursement rate, so good pricing supports a higher rate and more warranty gross. (6) Control the strategy: know how your matrix is configured and review it against your gross profit goals. Takeaway: the matrix is one of the most effective tools for improving profit without more volume.\n5 reasons parts gross profit % matters (Noah, Oct 26): (1) Maximizes profitability: higher GP % earns more from the same volume; small improvements add up over a year. (2) Measures pricing effectiveness: are parts sold at the right prices; where to improve matrices, discounts and markups. (3) Controls discounts and cost increases: shows when discounting cuts profit, and when rising parts costs are not reflected in selling prices. (4) Supports dealership profit: parts gross covers department expenses, adds to net profit and helps fixed ops absorb store overhead. (5) Identifies opportunities: compare against targets, history and other stores; find weak categories (customer pay, wholesale, warranty, internal). Takeaway: parts GP % measures how well we turn parts sales into profit; raising it is one of the most direct ways to grow profit without more volume.\n37 cents (Oct 26): P&L Detail, Jan-Sep 2026 same-store, parts department profit $22.75M / parts gross $61.93M = 36.7% (2025 37.1%); service 21.8% (2025 18.9%); new and used lose money at the department level; parts beats service in 31 of 33 stores (not Honda Henderson, VW Henderson). F&I income less F&I commissions is higher (about 80%) but F&I is not a P&L department and its other costs sit in new and used, so say 'department'. Department view: about $87M of other income (doc fees, packs, factory money, holdback) is booked outside the departments." });

X({ eyebrow: "Fixed Ops Profitability", title: "Absorption",
  formula: ["Fixed ops absorption", "=", "fixed ops gross", "÷", "total store fixed expenses"],
  define: "How much of the whole store's fixed expenses, not just service and parts expenses, the service and parts gross pays for.",
  example: [["At 100%", "fixed ops covers every fixed expense"], ["Group today", "78.1%"]], exampleLabel: "READING IT",
  why: ["At 100%, the store covers its fixed costs before it sells a car.",
        "It protects the store when car sales slow.",
        "It is the clearest measure of fixed ops' value to the dealership."],
  levers: [["Fixed ops gross", "Everything on the customer-pay slides: ROs, hours, ELR, margin"],
           ["Store fixed expenses", "The whole store's rent, facility and other fixed costs"],
           ["Warranty and internal", "Gross from warranty and reconditioning counts too"],
           ["Parts", "Counter and wholesale parts gross counts as well"]],
  ours: "Group: 78.1% absorption, up 2.1 pts. Fixed gross grew 4.1%; fixed expenses 1.3%.",
  tone: "g", notes: "PLACE: Profitability, in front of Fixed Ops Absorption %.\nP&L definition: line 254 Total Gross Profit - Fixed Ops / line 258 Fixed Expenses (the stored line 257 returns a dollar value, so the slide derives it). GMC Prescott includes its body shop, as on the P&L." });

X({ eyebrow: "Fixed Ops Profitability", title: "Net to Gross",
  formula: ["Net to gross", "=", "department profit", "÷", "department gross"],
  define: "How much of each gross dollar the department keeps after its own expenses.",
  why: ["It is the true measure of how the department performs.",
        "Gross is what we make; net is what we keep.",
        "Gross can grow while net falls if expenses grow faster."],
  // Oct 26 (Noah): pay plans once; semi-fixed examples; outrun expense; policy without "goodwill"
  levers: [["Personnel expense", "Headcount and pay plans: the biggest line"],
           ["Semi-fixed expenses", "Supplies, shop tools and equipment, training, advertising, outside services, data processing"],
           ["Policy and adjustments", "Service policy (work done at no charge to the customer) and warranty claims the manufacturer won't pay"],
           ["Gross", "Outrun your expenses: when net is low, it is either an expense problem or a gross problem"]],
  leverRows: [1, 1.35],
  ours: "Group: service 21.8%, up 2.9 pts; parts 36.7%, down 0.4 pts.",
  notes: "PLACE: Profitability, in front of Service Net to Gross %.\nP&L Detail statement: DEPARTMENT PROFIT OR LOSS / GROSS PROFIT / INCOME, Mechanical Dept (service) and P&A Dept (parts)." });

X({ eyebrow: "Opcodes", title: "Tracked Opcode Closing %",   // was What Is an Op Code Closing %? (Noah, Oct 25: too literal)
  formula: ["Closing %", "=", "CP ROs with the service", "÷", "all CP ROs"],
  define: "How often a customer-pay visit includes the service.",
  example: [["Alignment", "15%"], ["Tire balance", "15%"], ["Brake fluid", "10%"], ["Fuel service", "7.5%"]], exampleLabel: "BENCHMARKS",
  why: ["These cars are already in our service drives: the rate shows how well we use the opportunities we already have.",
        "Like labor dollars per RO, it is a direct reflection of the job we are doing on every car.",
        "Each point is worth thousands of services across the group."],
  levers: [["MPI on every car", "Find the need and show the customer"],
           ["Menu at write-up", "Time and mileage recommendations, every customer"],
           ["Consistency", "Present it every visit, no exceptions. Say it once and skip it the next time, and we lose credibility"],
           ["Tracking", "Daily Viz Op Code report; feedback before the 15th"]],
  leverRows: [1, 1.35],
  ours: "All four rose 1.2-1.4 pts and beat 2025 every month. Balance is furthest off, at 8.3%: 18,493 short of 15%.",
  notes: "PLACE: Opcodes, after the section break, in front of Alignment %.\nDenominator: CP ROs on the op code report basis (excluding ROs flagged out of the report)." });

X({ eyebrow: "Technician Proficiency", title: "Technician Proficiency",
  formula: ["Proficiency", "=", "hours sold (flagged)", "÷", "hours worked (clocked)"],
  define: "How many hours a technician produces for every hour they are on the clock. 85% is on target; 100% or more is above.",
  example: [["Works 40 hours", "sells 34"], ["Proficiency", "85%"], ["Sells 40", "100%"]], exampleLabel: "A 40-HOUR WEEK",
  // Oct 25 (Noah): service departments are production facilities; someone has to actively manage production
  why: ["A service department is a production facility. What we produce is hours.",
        "Any production floor knows what it made today and what each person made. So should we.",
        "It takes someone actively managing production, every day."],
  levers: [["Dispatch", "Work waiting at the bay; the right job to the right tech"],
           ["Car count", "Not enough work leaves techs waiting"],
           ["Parts staging", "Parts at the bay before the tech needs them"],
           ["Approvals", "Fast customer approval on the inspection"],
           ["Skill and training", "Tech level in line with job skill"],
           ["Proficiency vs efficiency", "Efficiency divides by time clocked on ROs, not all hours worked"]],
  ours: "Group: 94.7%. 20 of 33 stores at or above 85%; 262 of 583 counted techs below it.",
  tone: "n", notes: "PLACE: Technician Proficiency, after the section break, in front of the store ranker.\nProficiency uses ADP worked hours (regular + overtime), so it includes time not on an RO. Efficiency would use DMS time punched on jobs, which this data does not have." });

X({ eyebrow: "Technician Proficiency", title: "How We Measured It",
  formula: ["Sold hours (Qlik)", "÷", "worked hours (ADP, regular + overtime)"], formulaSize: 22,
  // Noah, Oct 26: label the period Jan-Sep 2026 YTD (pay periods run Dec 25, 2025 to Oct 9, 2026; in the notes)
  define: "January through September 2026, year to date. Store = total sold hours / total worked hours.",
  whyHead: "WHO COUNTS", leverHead: "THE RULES",
  why: ["Service technicians and express technicians",
        "Service team leaders and shop foremen",
        "Everyone whose job is to produce sold hours",
        "Not counted: body shop, detail and tint"],
  levers: [["Shop foremen above 30%", "Below that, a foreman is mainly dispatching"],
           ["Techs above 10%", "Below that, they are not working as a line tech"],
           ["Role changes", "Worked hours trimmed to the months a tech actually flagged"],
           ["Keying errors", "Example: 940 hours keyed on a single $0 line"],
           ["Shared tech numbers", "Example: hours flagged to a shop code like 999, not a person"],
           ["Opportunity", "Each tech under 85% raised to 85% of their own worked hours"]],
  ours: "The rules keep non-tech time and data errors out of the numbers, and every store is measured the same way.",
  tone: "n", notes: "PLACE: right after 'Technician Proficiency' (explainer).\nPeriod: pay periods Dec 25, 2025 to Oct 9, 2026, matched by tech number, all pay types; labelled Jan-Sep 2026 YTD on the slide (Noah, Oct 26: closest the pay periods allow).\nRules as applied on the proficiency slides (Oct 16 2026). Express techs at stores that pool express hours are mostly excluded by the 10% rule." });

// 13. Gross section (Noah, Oct 22): how service and parts gross are built, in front of Total Gross Profit - Fixed Ops
X({ eyebrow: "Fixed Ops Gross", title: "Fixed Ops Gross",
  formula: ["Fixed ops gross", "=", "service gross", "+", "parts gross"],
  define: "What each department keeps after the cost of what it sold: labor sale less technician cost, parts sale less parts cost.",
  example: [["Service", "$126.87M"], ["Parts", "$61.93M"], ["Fixed ops", "$188.80M"]], exampleLabel: "JAN–SEP 2026",
  whyHead: "HOW EACH IS BUILT",
  why: ["Service earns only on labor hours sold: the labor sale less what the technician is paid for the job.",
        "Parts sold on a service RO are parts gross, not service gross.",
        "Both include customer pay, warranty and internal work."],
  // Oct 25 (Noah): process on the labor sale, right work right tech, parts follow labor, wholesale has to pay for itself
  levers: [["Service: labor sale", "A written process, followed on every car, every time, tracked and managed daily"],
           ["Service: technician cost", "Right work, right tech. Hourly pay, guarantees and bonuses paid off the RO are unapplied labor and lower gross"],
           ["Parts: on the RO", "Every dollar of labor sold brings more parts gross with it"],
           ["Parts: over the counter", "Retail, wholesale, tires and oil. Wholesale has to make sense: its gross must cover what it costs to run"]],
  leverRows: [1.6, 1.3],
  ours: "Group: $188.80M fixed ops gross Jan–Sep 2026, up $7.47M (4.1%): service +$6.30M, parts +$1.17M.",
  tone: "g", notes: "PLACE: Gross section, after Service Is Sales 2.0 and in front of Total Gross Profit - Fixed Ops.\nSource: Qlik Profit & Loss, Total Gross Profit - Service and Total Gross Profit - Parts, same-store, INEOS excluded, GMC Prescott body shop gross excluded (as on the gross slides). Hourly technician pay, guarantees and bonuses paid off the RO post as unapplied labor: not charged on any repair order, a cost-of-sale adjustment that reduces service gross (Noah, Oct 23)." });

// 14. Oct 27 (Noah): the three technician measures side by side, after the Technician Proficiency explainer.
// Revised Oct 27: "hours flagged" throughout, available hours vs hours clocked to repair orders kept apart,
// one labelled example (8 available, 6 clocked to ROs, 7.5 flagged).
{
  const { C, F, sz } = H, s = pptx.addSlide();
  const TAKE = "Productivity measures time spent on repair orders; efficiency measures hours flagged against time on repair orders; proficiency measures total hours flagged against available time.";
  H.page(pptx, s, { eyebrow: "Technician Proficiency", title: "Proficiency, Efficiency and Productivity" });
  H.panel(pptx, s, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(TAKE), label: null });
  const t = (txt, o) => s.addText(H.words(txt, "definitions"), { fontFace: F, isTextBox: true, margin: 0, ...o });
  const tiles = [
    ["Productivity", "Hours clocked to repair orders ÷ available hours", "How much of the technician's available time was spent on repair orders?", "Example: 6 ÷ 8 = 75%", "Not measured in our data"],
    ["Efficiency", "Hours flagged ÷ hours clocked to repair orders", "How many labor hours did the technician produce for each hour clocked to repair orders?", "Example: 7.5 ÷ 6 = 125%", "Not measured in our data"],
    ["Proficiency", "Hours flagged ÷ available hours", "How many labor hours did the technician produce compared with their total available time?", "Example: 7.5 ÷ 8 = 93.75%", "Target 85% · group 94.7%"]];
  const gap = 0.22, X0 = 0.78, W = (11.77 - gap * 2) / 3, Y = 1.55;
  tiles.forEach(([lab, f, q, ex, st], i) => {
    const x = X0 + i * (W + gap), mine = i === 2;
    s.addShape(pptx.ShapeType.rect, { x, y: Y, w: W, h: 3.2, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y: Y, w: W, h: 0.07, fill: { color: C.blue } });
    t(lab.toUpperCase(), { x: x + 0.2, y: Y + 0.2, w: W - 0.4, h: 0.3, fontSize: sz(14, "tile label"), bold: true, color: C.ink600, charSpacing: 0.8 });
    t(f, { x: x + 0.2, y: Y + 0.55, w: W - 0.4, h: 0.75, fontSize: sz(18, "definition formula"), bold: true, color: C.navy, valign: "top" });
    t(q, { x: x + 0.2, y: Y + 1.38, w: W - 0.4, h: 0.95, fontSize: sz(15, "definition question"), color: C.ink, valign: "top" });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: Y + 2.37, w: W - 0.4, h: 0.012, fill: { color: C.rule } });
    t(ex, { x: x + 0.2, y: Y + 2.45, w: W - 0.4, h: 0.34, fontSize: sz(16, "definition example"), bold: true, color: C.navy, valign: "middle" });
    t(st, { x: x + 0.2, y: Y + 2.8, w: W - 0.4, h: 0.3, fontSize: sz(14, "definition status"), bold: mine, color: mine ? C.ink : C.ink600, valign: "middle" });
  });
  // the one labelled example, and the identity that ties the three together
  const EY = Y + 3.2 + 0.2;
  s.addShape(pptx.ShapeType.rect, { x: X0, y: EY, w: 11.77, h: 1.2, fill: { color: C.blueTint } });
  t("ILLUSTRATIVE EXAMPLE", { x: X0 + 0.2, y: EY + 0.14, w: 11.37, h: 0.28, fontSize: sz(14, "example label"), bold: true, color: C.ink600, charSpacing: 0.8 });
  t("A technician has 8 available hours, clocks 6 hours to repair orders and flags 7.5 hours.", { x: X0 + 0.2, y: EY + 0.46, w: 11.37, h: 0.32, fontSize: sz(16, "example"), color: C.ink, valign: "middle" });
  t("Productivity × efficiency = proficiency:  75% × 125% = 93.75%", { x: X0 + 0.2, y: EY + 0.8, w: 11.37, h: 0.32, fontSize: sz(16, "identity"), bold: true, color: C.navy, valign: "middle" });
  H.takeaway(pptx, s, TAKE, "n");
  H.foot(pptx, s, "Standard service-department definitions. On the store slides, proficiency = Qlik hours flagged (sold hours) ÷ ADP regular + overtime hours, which stand in for available hours. Efficiency and productivity need DMS time clocked to repair orders, which this data does not include, so they are defined here but not measured. The example is illustrative, not a store or technician.");
  s.addNotes("PLACE: right after the Technician Proficiency explainer, before How We Measured It (Noah, Oct 27).\nTalk track: walk the example left to right. The tech was on repair orders 6 of 8 hours (75% productive) and flagged 7.5 hours in those 6 (125% efficient), so 7.5 flagged hours for 8 available (93.75% proficient). Proficiency is the one we measure: 85% target, group 94.7%. 'Productivity' here is the technician measure, not the Fixed Ops Productivity section (gross and ROs per employee).");
}

await pptx.writeFile({ fileName: "Metric Explainers - Review.pptx" }); console.log("built", pptx._slides.length);
