/* Talking-point slides, Oct 29 day 2 (built 2026-10-19). "What we said" lines come from the Service Manager
   Roundtable (Nov 15, 2024) and the Jan and May 2026 GM decks (project files). "Where we are" numbers are the
   verified figures already on the review deck slides (same-store, 33 stores, Jan-Sep 2026 vs 2025) and the
   technician proficiency slides. Each slide's notes say where it goes in the master deck and where each line came from. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const RT = "Roundtable 2024", GM = "GM meetings Jan & May 2026", JAN = "GM meeting Jan 2026", BOTH = "Roundtable & GM meetings";
const SRC = "Talking points: Service Manager Roundtable (Nov 15, 2024) and GM meeting decks (Jan and May 2026). Data: Qlik (Profit & Loss, Closed Repair Orders, Productivity), Jan-Sep 2026 vs 2025, same-store, 33 stores, INEOS excluded.";
// saidFrom stays in each call as the record of where the line came from; the slide no longer shows it (Noah, Oct 21)
const T = (o) => H.talkingPoints(pptx, pptx, { source: SRC, ...o, saidFrom: null });

// Why We Are Here (Noah, Oct 24): what fixed ops is worth to the whole group, GM statement net profit, all stores.
const D = JSON.parse(fs.readFileSync("./pulls/net_by_dept.json", "utf8")).all_stores;   // whole group, INEOS in (Noah, Oct 24)
const fx = y => -(D[y].service + D[y].parts), net = y => -D[y].net_gm;
const M = v => "$" + (v / 1e6).toFixed(2) + "M", usd = v => "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
const share = y => fx(y) / net(y) * 100, dFx = fx("y26") - fx("y25"), dNet = net("y26") - net("y25");
H.talkingPoints(pptx, pptx, { source: "Qlik - Profit & Loss, Jan-Sep 2026 vs 2025, all stores (INEOS included, to match the group P&L). Fixed ops = service (Mechanical) + parts (P&A) department profit: department gross less the department's own expenses. Group net = Net Profit Or Loss on the GM statement.",
  eyebrow: "Fixed Ops", title: "Why We Are Here", saidFrom: null,
  said: ["Foster collaboration, introduce ideas that drive operational excellence, and set a clear path to exceptional service performance",
         "Focus on profit-generating activities: the work that brings in incremental dollars, not just the work that fills the day",
         "Work on the business, not just in it: step out of the day-to-day long enough to fix the process behind the problems",
         "No risk of judgement: your success is our success",
         "Customer first, company second, team third",
         "You get what you tolerate",
         "Leverage the power in this room: collaboration"],
  now: [{ big: `${share("y26").toFixed(0)}%`, label: `of the group's net profit came from service and parts: ${M(fx("y26"))} of ${M(net("y26"))} (${share("y25").toFixed(0)}% in 2025).`, tone: "g" },
        // Oct 26 (Noah): the $ tile is out; how much of each gross dollar each side keeps (Pay Yourself First data)
        (() => { const K = JSON.parse(fs.readFileSync("./pulls/dept_keep.json", "utf8")), f = K.fixed, v = K.sales;
          const fk = (f.profit_service + f.profit_parts) / (f.gross_service + f.gross_parts);
          const vk = (v.profit_new + v.profit_used + v.gross_fi - v.fi_comp) / (v.gross_new + v.gross_used + v.gross_fi);
          return { big: `${Math.round(fk * 100)}\u00a2 vs ${Math.round(vk * 100)}\u00a2`, label: "Of every gross dollar, what service and parts keep after their own expenses vs new, used and F&I.", tone: "g" }; })(),
        { big: "+" + usd(dFx), label: `More fixed ops profit than 2025, while group net profit ${dNet < 0 ? "fell" : "rose"} ${M(Math.abs(dNet))}.`, tone: "g" }],
  ask: "leave today with one or two things you will put in place before year end, and tell the room what they are.",
  notes: `PLACE: right after the agenda, before the Gross section. In the deck Oct 24 (Noah): group net profit view replaces the gross / absorption / net-to-gross tiles. Group-level view of fixed ops' weight in the dealership's profit, Jan-Sep. Fixed ops department profit ${usd(fx("y26"))} (2025 ${usd(fx("y25"))}); group net profit ${usd(net("y26"))} (2025 ${usd(net("y25"))}). On the P&L, F&I income, other income and store-level costs are not split by department, so 'the rest of the store' is the group net less fixed ops department profit. Group net = Net Profit Or Loss on the GM statement (P&L summary), all stores; INEOS (opened 2026) is in both lines. Body shop (GMC Prescott) department profit is left out of fixed ops, as everywhere in the deck.` });

T({ eyebrow: "Fixed Ops Gross", title: "Service Is Sales 2.0", saidFrom: BOTH,
  said: ["Sell hours, build value, sell peace of mind",
         "Drive visits: don't say no to dollars",
         "Generate throughput: max out hours",
         "Drive retention and sell cars",
         "Sales sells the first car. Service sells the next six.",
         "Every car, every time: follow the processes you have in place, consistently",
         "Processes only work when they are actively managed"],   // Noah, Oct 26
  now: [{ big: "+$6,297,386", label: "Service gross, up 5.2%, ahead of 2025 in all nine months.", tone: "g" },
        { big: "+$1,171,181", label: "Parts gross, up 1.9%, ahead of 2025 in six of nine months.", tone: "g" },   // same pattern as service (Noah, Oct 24)
        { big: "$20.59M", label: "Fixed ops gross a month over the last twelve months.", tone: "g" }],
  ask: "run service like a sales department: every visit is a chance to sell the next hour and the next car.",
  notes: "PLACE: Gross section, after the section break and before Total Gross Profit - Fixed Ops.\nSaid: Roundtable slide 9 ('Service Department?'); Jan GM slide 58 (Sales and Service Relationship).\nFor the GMs who come from sales (Noah, Oct 25): service and parts keep far more of every gross dollar than vehicle sales does. Jan-Sep 2026, after each department's own expenses, fixed ops kept 27 cents of every gross dollar; new, used and F&I together kept 7 cents (before doc fees, packs and factory money, which the P&L books outside the departments). People cost 45 cents of every fixed ops gross dollar and 62 cents of every vehicle sales gross dollar. That is why paying yourself first, on internal work, matters: see Pay Yourself First at the start of the Internal section." });

T({ eyebrow: "Customer Pay", title: "What Can You Control?", saidFrom: RT,
  // Oct 26 (Noah): same points, smoother order; track daily; take care of the customers we have
  said: ["Take care of the customers we already have: how we treat them decides whether they come back",
         "Repair order count: how many customers we serve",
         "Dollars per repair order: how much we sell each one",
         "The rates we charge, and how we decide what to charge",
         "CP labor $ per RO is the best way to judge performance, store and individual",
         "Track the metrics that matter most every day, talk about them, and train on them"],
  now: [{ big: "279,756 ROs", label: "Customer-pay RO count, up only 456. Volume is flat.", tone: "n" },
        { big: "1.87 hrs per RO", label: "Up 0.07, ahead in all nine months. Every 0.1 hour per RO is worth $4,543,691 of labor.", tone: "g" },
        { big: "$304 per RO", label: "CP labor per RO, up $17. The $4,899,376 labor gain came from selling more per visit.", tone: "g" }],
  ask: "with car count flat, the growth is in hours per RO. Sell the next tenth of an hour on every RO.",
  tone: "g", notes: "PLACE: CP Engine section, after the section break and before CP RO Count.\nSaid: Roundtable slides 18-19 (KPIs; CP Labor $ per RO)." });

T({ eyebrow: "Customer Pay", title: "The Front Door", saidFrom: GM,
  said: ["Open and available scheduler: same-day or next-day appointments, quick online scheduling",
         "Handle basic maintenance quickly and efficiently. Saturdays are what you choose to make them.",
         "Every call is a sales opportunity: a mishandled call costs us the customer",   // after-hours / Sundays point dropped (Noah, Oct 26)
         "No-show process; first appointment at time of sale; next appointment before the customer leaves",
         "Active retention tracking: work the Viz 'Close to Losing' and 'Unretained' lists"],
  now: [{ big: "+0.2%", label: "Customer-pay RO count. 13 of 33 stores had fewer CP ROs than last year.", tone: "n" },
        { big: "$4,473,284", label: "A year in labor and parts gross from one more customer-pay RO a day at every store.", tone: "n" }],   // ROs per advisor moved to Open Decisions (Oct 24)
  ask: "every customer leaves with the next appointment booked, and every call reaches someone who asks for it.",
  notes: "PLACE: CP Engine section, after the CP RO Count slides and before Hours per RO.\nSaid: Jan/May GM slides on Appointment Scheduling, Appointment Discipline, Phone Call Handling and Retention Management (Jan 2, 7, 9, 11; May 2, 4, 8, 9).\nOptional: the Flai voice AI service view (Jul 16-Aug 15 pull: 4,368 of 8,351 transfers rang out) fits here, but it is a different source and period; refresh it before use and keep it off any slide with a P&L number." });

T({ eyebrow: "Customer Pay · Effective Labor Rate", title: "Pricing Strategy", saidFrom: BOTH,   // marks it as part of the ELR run (Noah, Oct 26)
  said: ["Do a comprehensive market labor study",
         "Calculate your ELR on every maintenance item; target ELR by work type",
         "Know your loss leader, and what the competition is charging",
         "No direct price changes to parts or labor: use ADC-controlled discounting for every discount",
         "Benchmark: CP ELR at least 70% of door rate"],
  now: [{ big: "$162.42", label: "CP effective labor rate, up $3.56 and ahead in all nine months.", tone: "g" },
        { big: "74.7%", label: "CP ELR as a share of door rate, Jul-Sep. 25 of 33 stores at or above 70%.", tone: "g" },
        { big: "$127K a month", label: "Added labor if the 8 stores under 70% reached it on the hours they already sell, Jul-Sep: $126,831 a month.", tone: "n" }],   // opportunity size: blue
  ask: "know what you collect on every op code, and run every discount through ADC so we can see it.",
  notes: "PLACE: CP Engine section, before the CP ELR slides (ahead of the door-rate options).\nSaid: Roundtable slide 15 (Pricing Strategy); Jan/May GM Road to the Sale slide (pricing lines); Jan GM slide 54 (Discounting Discipline); 70% benchmark from the Jan/May door rate slides.\nOct 23 (Noah): grid vs non-grid tile replaced. $126,831 a month = Jul-Sep, for each of the 8 stores under 70% (Chevy LV, Toyota Spokane, Cadillac, Kia LV, Subaru Prescott, Lincoln, GMC Prescott, Hyundai Prescott): (70% x door rate - CP ELR) x CP hours sold, summed ($380,494) / 3. Door rates from the store door rate sheet (door.json), ELR and hours from elr_q3.json." });

T({ eyebrow: "Customer Pay", title: "The Goldilocks Margin", saidFrom: RT,
  said: ["Gross profit % is the share of labor you keep after paying the tech",
         "Not too high, not too low: 78% to 80%",
         "Skill code costing: a hybrid of traditional and percentage",
         "Right tech, right job"],
  now: [{ big: "80.8%", label: "CP labor gross profit, down 0.1 pts. Just above the 78-80% target.", tone: "n" },
        { big: "40.7%", label: "CP parts gross profit, up 1.2 pts and ahead in all nine months.", tone: "g" }],
  ask: "hold labor margin in the 78-80% band by matching the tech to the job, not by squeezing tech pay.",
  notes: "PLACE: CP Engine section, before CP Labor Gross Profit %.\nSaid: Roundtable slide 21 (Gross Profit %). Above 80% can mean techs are underpaid on the work they do; below 78% means the job mix or the dispatch is off." });

// Oct 25 (Noah): was "Express Service". The point is capitalizing on the maintenance visits we already have.
T({ eyebrow: "Opcodes", title: "Basic Maintenance Service", saidFrom: GM,
  // Oct 26 (Noah): "the value of the unknown" moved from the right column into these two bullets
  said: ["On-demand maintenance appointments: get cars in quickly. A car we can't book is work we never see",
         "A complete, comprehensive multi-point inspection on every visit: what it finds is worth far more than the maintenance the car is due for",
         "Experienced eyes on every car: entry-level staff alone won't know what to look for",
         "An oil change that misses worn brakes or a leaking water pump is a disservice to the customer",
         "More cars in the door is great, but first make the most of every car already in the shop",
         "Processes and policies that make it happen: every car, every time"],
  now: [{ big: "54.3%", label: "Maintenance share of customer-pay labor across the group; repair 30.5%. Depending on work mix, maintenance runs 25% to 75% of a store's service.", tone: "n" },
        // Oct 26: MPI completion replaces "The value of the unknown" (now in the bullets); same figure as Tracking & Accountability
        { big: "81.3%", label: "MPI completion, July to September (store-reported). 12 of 33 stores at the 90% benchmark: every car skipped is unknown work left in the car.", tone: "b" }],
  ask: "capitalize on every maintenance visit: in quickly, inspected completely by someone who knows what to look for, and every finding recommended.",
  notes: "PLACE: Opcode Benchmarks, after Road to the Sale and ahead of MPI Completion % (Noah, Oct 23: maintenance visits and a complete MPI go hand in hand).\nTalk track (Noah, Oct 25): go to the dentist for a quick cleaning and they don't tell you something is starting to look like a cavity: that is a disservice. Get cars in quickly, perform a complete, comprehensive MPI, and make the proper recommendations at that time. This is critical, and it can't be left to entry-level employees alone: someone with experience who knows what to look for has to be looking at these cars. Depending on work mix, maintenance can be 25% to 75% of a store's service. Many stores focus on getting more cars in the door, which is great, and we want both, but they miss the opportunities they already have. Processes and policies make sure we capitalize on them: every car, every time.\nThe 52.5% customer-pay share of labor sale tile was dropped with the rewrite.\nSaid: Jan GM slide 8 / May GM slide 7 (Express Service), reworded." });

T({ eyebrow: "Opcodes", title: "Tracking & Accountability", saidFrom: BOTH,
  said: ["Daily review of the Viz 'Op Code' and 'Performance Report' and the Fixed Ops DOC",
         "Give feedback quickly and often. Trends start early: don't wait until the 15th.",
         "Tracked metrics improve performance; we are responsible for the knowing",
         "Scoreboard techs and advisors: CP $ per RO, ELR, MPI completion, closing %, key op codes"],
  // 15.2% group figure removed with the Unclassified Customer-Pay Labor slide (Noah, Oct 22); the range now explains itself
  now: [{ big: "81.3%", label: "MPI completion, July to September (store-reported). 12 of 33 stores at the 90% benchmark.", tone: "b" }],
  ask: "if it isn't described, it can't be tracked. Clean up op code descriptions so the reports show what you sell.",
  notes: "PLACE: Opcode Benchmarks, after MPI and Video Completion %, ahead of the op code closing % slides (moved out of Internal, Oct 23).\nSaid: Jan GM slide 19 / May GM slide 15 (Tracking & Accountability); Roundtable slide 16 (Performance Tracking)." });

T({ eyebrow: "Opcodes", title: "Road To The Sale", saidFrom: BOTH,
  said: ["Every advisor, every customer: timely meet and greet at the vehicle",
         "Walk-around with the customer: hood, tires and wipers",
         "Review history and previously declined items",
         "Time and mileage recommendations from an accurate, available menu",
         "MPI mandate: every car, every time",
         "Consistency is key: trained, documented and actively managed",
         "We produce hours. Every service or repair not recommended is hours a technician never gets the chance to earn"],
  now: [{ big: "Up in all four", label: "Alignment, balance, brake fluid and fuel service each rose 1.2-1.4 pts and beat 2025 every month.", tone: "g" },
        { big: "18,493", label: "Balances short of the 15% benchmark. The group is at 8.3%.", tone: "b" },
        { big: "4,821", label: "Alignments short of 15%; the group is at 13.2%. Brake fluid is 0.7 pts from its 10%.", tone: "b" }],
  ask: "every car gets the walk-around and the menu. The benchmarks are where that process lands.",
  notes: "PLACE: Opcodes section, after the section break and before Alignment %.\nSaid: Roundtable slides 11, 12, 14 (Road to the Sale, Maintenance Menus, MPI Mandate); Jan GM slide 15 / May GM slide 11 (Road to the Sale). Fuel service is 5.8% vs 7.5% (4,731 short)." });

T({ eyebrow: "Warranty", title: "Warranty Compliance", saidFrom: JAN,
  said: ["Higher labor rates, higher parts reimbursement and tariffs are driving up OEM scrutiny",
         "Six audits cost over $400,000",
         "Primary causes: written documentation, manager signatures on add-ons and in-stock units, out mileage, punch times, manufacturer authorization",
         "The general manager is responsible for warranty policy being followed",
         "Spot-check repair orders and self-audit"],
  now: [{ big: "$3,788,200", label: "Added gross profit from warranty labor and parts rate increases approved this year, 19 stores.", tone: "g" },
        { big: "$197.17", label: "Warranty ELR, up $16.93; 30 of 33 stores raised it.", tone: "g" },
        { big: "44.7%", label: "Warranty parts gross profit, up 2.1 pts and ahead of 2025 in all nine months.", tone: "g" }],
  ask: "a higher warranty rate brings more scrutiny. Self-audit before the factory does.",
  notes: "PLACE: Warranty section, after the section break, ahead of Warranty Rate Increases (Oct 22: $3,788,200 added gross profit from the YTD RWR Financials vendor report replaced the warranty RO count here).\nSaid: Jan GM slides 13-14 (audit chargebacks; Warranty Compliance). The Jan slide named six stores' chargebacks; kept to the $400,000 total here, in line with no store call-outs from the podium." });

T({ eyebrow: "Fixed Ops Productivity", title: "People", saidFrom: BOTH,
  said: ["Our most valuable asset: this is what separates us",
         "Invest in training and a defined career path. Be the best place to work.",
         "Grow your own techs, advisors and managers: apprentice programs, develop your bench, train your replacement",
         "Always be teaching, training and educating",
         "Being short-handed limits scheduling, loses opportunities and leads to missed MPIs",
         "Clear pay plans, job descriptions and expectations"],
  now: [{ big: "−76", label: "Fixed ops staff, Sep to Sep, to 1,540 (−4.7%): technicians −17, advisors −7, managers −3.", tone: "n" },
        { big: "$13,617", label: "Fixed gross per fixed employee a month, up $1,207. Gross grew 4.1% with 5.1% fewer people.", tone: "g" },
        // Noah, Oct 26: what an empty bay misses. prof/opp.json: 841,721 sold hrs / 583 counted techs / 9.49 months x $187.98 gross per sold hour
        { big: "$28,600", label: "Gross an average tech's bay produces a month: 152 sold hours at $188 each. An empty or idle bay misses all of it.", tone: "b" }],
  ask: "keep every seat filled with someone who adds hours. Being short-handed costs appointments, MPIs and sold hours.",
  notes: "PLACE: Productivity section, after the section break and before Fixed Gross per Fixed Employee.\nSaid: Roundtable slide 6 (People); Jan GM slide 57 (Technician Pipeline).\nBay tile (Oct 26): counted technicians (proficiency basis, Dec 25 2025 - Oct 9 2026, 9.49 months) flagged 841,721 sold hours, 583 techs = 152 sold hours per tech a month; labor gross $132.13 + parts gross $55.85 = $187.98 per sold hour, all pay types (Qlik Closed ROs) = $28,584 a month per tech. One tech per bay assumed (1:1 tech to stall). The -2.3% CP & W RO tile was dropped for it.\nNumbers rebuilt 2026-10-20 on ADP headcount (Employee Master extract 10/19/2026); the Qlik headcount history undercounted 2025." });

T({ eyebrow: "Fixed Ops Productivity", title: "Pay Plans Are Work Instructions", saidFrom: GM,
  said: ["Clear and easily understood",
         "Based on customer-pay performance: CP labor or gross per RO, hours per RO, maintenance penetration, total sales or gross",
         "Incentivize improvement in the metrics that make the biggest difference",
         "Move technicians to flat rate and away from hourly: it rewards productivity and makes proficiency easy to track",
         "Advisor pay: target a 50/50 split of base commission and performance bonus, with about half of earning potential tied to measurable metrics",
         "Every system is perfectly designed to get the results it gets"],
  now: [{ big: "$304", label: "CP labor per RO, the number the plans should pay on. Up $17.", tone: "g" },
        { big: "94.7%", label: "Technician proficiency, Jan-Sep: 20 of 33 stores at or above 85%. Flat-rate pay rewards what techs produce.", tone: "g" }],   // replaces ROs per advisor, now in Open Decisions (Noah, Oct 23)
  ask: "check that every advisor and tech pay plan pays for the numbers on these slides.",
  notes: "PLACE: Productivity section, before CP & W ROs per Service Advisor.\nSaid: Jan GM slide 16 / May GM slide 12 (Pay Plans aka Work Instructions)." });

T({ eyebrow: "Technician Proficiency", title: "Technicians & Stalls", saidFrom: BOTH,
  said: ["1:1 tech to stall unless you run on-demand appointments",
         "Tech level in line with job skill: right tech, right job",
         "5/6 of staff on schedule at all times, Saturdays included",
         "90%+ MPI completion; track and report technician efficiency",
         "Use scoreboarding and communicate to techs where they stand",
         "Set clear objectives and give feedback: no one can be held to a goal they don't know. We're responsible for the knowing"],   // Noah, Oct 25
  now: [{ big: "94.7%", label: "Group technician proficiency (sold hours / worked hours). 20 of 33 stores at or above 85%.", tone: "g" },
        { big: "262 of 583", label: "Counted technicians under 85%. Every store has at least one.", tone: "b" },
        { big: "$1.65M a month", label: "Gross missed if each of them reached 85%: no extra hours, techs or bays.", tone: "b" }],
  ask: "set clear goals, show every tech where they stand every week, and give feedback. The best techs in each store already prove 85% is possible.",
  notes: "PLACE: Technician Proficiency section, after the section break and before the store proficiency ranker.\nSaid: Roundtable slide 22 (Technicians & Stalls); Jan GM slide 52 / May GM slide 16 (MPI Completion and Tech Efficiency).\nNow: proficiency slides (ADP worked hours + Qlik sold hours, Dec 25-Oct 9, with the foreman 30%, tech 10%, role-change and keying-error rules)." });

T({ eyebrow: "Fixed Ops", title: "Before You Leave", saidFrom: RT,
  said: ["What did you learn today?",
         "Give us one or two things you are going to act on and implement",
         // Oct 27 (Noah): "Do you want to do this again?" dropped, this is a recurring meeting
         "Pick the one number you will check every day, and name who owns it",
         "Call a store that leads where you trail: ask what they do differently",
         "Reach this room in FMGSVCMGR",
         // Oct 27 (Noah): commitments get reviewed
         "We will review your commitments at the next meeting"],
  nowLabel: "Where the money is",
  now: [{ big: "0.1 hr per RO", label: "Worth $4,543,691 of labor at today's ELR.", tone: "g" },
        // Oct 27 (Noah): replaces the 18,493 balances tile; total from opmiss-deck.mjs (Every Store at Benchmark)
        { big: "$4.9M a year", label: "Every store at all four opcode benchmarks: $4,895,199 in labor and parts gross.", tone: "b" },
        { big: "$1.65M a month", label: "Technician proficiency, lifting each tech under 85% to 85%.", tone: "b" }],
  ask: "name your one or two commitments out loud, with a number attached, before you leave the room.",
  tone: "g", notes: "PLACE: the close, after the last Productivity / Proficiency slide.\nSaid: Roundtable slide 24 (Misc & Closing). Commitments are reviewed at the next meeting (Noah, Oct 27): keep a list of what each store names." });

// Pay Yourself First (Noah, Oct 25): opens the Internal section. Internal work isn't service making money off sales;
// it keeps gross in the department that holds on to the most of it. Data: pulls/dept_keep.json (P&L Detail, Jan-Sep 2026).
{
  const D = JSON.parse(fs.readFileSync("./pulls/dept_keep.json", "utf8")), f = D.fixed, s = D.sales;
  const fg = f.gross_service + f.gross_parts, fp = f.profit_service + f.profit_parts, fpe = f.personnel_service + f.personnel_parts;
  const sg = s.gross_new + s.gross_used + s.gross_fi, sp = s.profit_new + s.profit_used + s.gross_fi - s.fi_comp;
  const spe = s.personnel_new + s.personnel_used + s.slsp_comp_new + s.slsp_comp_used + s.fi_comp;
  const c = v => Math.round(v * 100) + "\u00a2", M = v => "$" + (v / 1e6).toFixed(1) + "M";
  T({ eyebrow: "Internal", title: "Pay Yourself First", saidFrom: null,
    said: ["Internal work isn't service making money off sales: it is gross the store earns on its own cars",
           "Sales prices a car off what it costs, within reason. A fair internal charge is part of that cost",
           "Service and parts keep far more of every gross dollar than vehicle sales does",
           "Every gross dollar in sales carries commissions, pay and expense that fixed ops doesn't",
           "Price internal labor and parts like customer pay: pay yourself first"],
    now: [{ big: c(fp / fg), label: `Fixed ops keeps of every gross dollar after its own expenses: ${M(fp)} of ${M(fg)}.`, tone: "g" },
          { big: c(sp / sg), label: `New, used and F&I keep of every gross dollar: ${M(sp)} of ${M(sg)}, before other income such as doc fees, packs and factory money.`, tone: "n" },
          { big: `${c(fpe / fg)} vs ${c(spe / sg)}`, label: "Of every gross dollar spent on people and commissions: fixed ops vs vehicle sales.", tone: "n" }],
    ask: "hold internal labor and parts margins to customer pay. Every internal dollar is gross in the department that keeps the most of it.",
    source: "Qlik - Profit & Loss, Detail statement, Jan-Sep 2026, same-store. Kept = department gross less the department's own expenses. Vehicle sales = new + used departments plus F&I income and F&I compensation. Other income not assigned to a department (doc fees, packs, factory money, holdback: $87.2M net) and management fees are left out of both.",
    notes: `PLACE: Internal section, after the section break and before Internal ELR (Noah, Oct 25).\nTalk track: a lot of stores see internal as service making money off sales. It isn't: sales prices a car off its cost, within reason, so a fair internal charge is part of the deal, and the gross lands in the department that keeps the most of it. Example: if sales makes $500 of gross on a car, the commissions, pay and expense that come out of it leave far less than the same $500 of internal gross keeps in service and parts.\nNumbers (P&L Detail, Jan-Sep 2026, same-store): fixed ops gross ${M(fg)}, department profit ${M(fp)} (${c(fp / fg)}); people ${M(fpe)} (${c(fpe / fg)}). Vehicle sales gross (new ${M(s.gross_new)}, used ${M(s.gross_used)}, F&I ${M(s.gross_fi)}) ${M(sg)}; new and used department results ${M(s.profit_new)} and ${M(s.profit_used)}, less F&I compensation ${M(s.fi_comp)}: ${M(sp)} kept (${c(sp / sg)}); people and commissions ${M(spe)} (${c(spe / sg)}). Caveat: $87.2M of net other income (doc fees, packs, factory money, holdback, interest) is booked outside the departments and most of it comes with vehicle sales; with it, vehicle sales keeps much more. The comparison is department against department, as the P&L reports them.` });
}

await pptx.writeFile({ fileName: "Talking Points - Day 2.pptx" }); console.log("tp built", pptx._slides.length);
