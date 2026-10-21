/* Talking-point slides, Oct 29 day 2 (built 2026-10-19). "What we said" lines come from the Service Manager
   Roundtable (Nov 15, 2024) and the Jan and May 2026 GM decks (project files). "Where we are" numbers are the
   verified figures already on the review deck slides (same-store, 33 stores, Jan-Sep 2026 vs 2025) and the
   technician proficiency slides. Each slide's notes say where it goes in the master deck and where each line came from. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const RT = "Roundtable 2024", GM = "GM meetings Jan & May 2026", JAN = "GM meeting Jan 2026", BOTH = "Roundtable & GM meetings";
const SRC = "Talking points: Service Manager Roundtable (Nov 15, 2024) and GM meeting decks (Jan and May 2026). Data: Qlik (Profit & Loss, Closed Repair Orders, Productivity), Jan-Sep 2026 vs 2025, same-store, 33 stores, INEOS excluded.";
// saidFrom stays in each call as the record of where the line came from; the slide no longer shows it (Noah, Oct 21)
const T = (o) => H.talkingPoints(pptx, pptx, { source: SRC, ...o, saidFrom: null });

T({ eyebrow: "Fixed Ops", title: "Why We Are Here", saidFrom: RT,
  said: ["Foster collaboration, introduce ideas that drive operational excellence, and set a clear path to exceptional service performance",
         "No risk of judgement: your success is our success",
         "Customer first, company second, team third",
         "You get what you tolerate",
         "Leverage the power in this room: collaboration"],
  now: [{ big: "+$7,468,567", label: "Fixed ops gross, up 4.1% year to date. 25 of 33 stores grew.", tone: "g" },
        { big: "78.1%", label: "Fixed ops absorption, up 2.1 pts. Fixed gross grew 4.1%, fixed expenses 1.3%.", tone: "g" },
        { big: "21.8%", label: "Service net to gross, up 2.9 pts and ahead of 2025 in all nine months.", tone: "g" }],
  ask: "leave today with one or two things you will put in place before year end, and tell the room what they are.",
  tone: "g", notes: "PLACE: right after the cover, before the Gross section.\nSaid: Roundtable slides 2, 5, 7, 6 (objective; philosophy & culture; 'you get what you tolerate'; collaboration). Ask adapted from Roundtable slide 24 closing questions." });

T({ eyebrow: "Fixed Ops Gross", title: "Service Is Sales 2.0", saidFrom: BOTH,
  said: ["Sell hours, build value, sell peace of mind",
         "Drive visits: don't say no to dollars",
         "Generate throughput: max out hours",
         "Drive retention and sell cars",
         "Sales sells the first car. Service sells the next six."],
  now: [{ big: "+$6,297,386", label: "Service gross, up 5.2%, ahead of 2025 in all nine months.", tone: "g" },
        { big: "+$1,171,181", label: "Parts gross, up 1.9%, but parts net to gross slipped 0.4 pts to 36.7%.", tone: "n" },
        { big: "$20.59M", label: "Fixed ops gross a month over the last twelve months.", tone: "g" }],
  ask: "run service like a sales department: every visit is a chance to sell the next hour and the next car.",
  notes: "PLACE: Gross section, after the section break and before Total Gross Profit - Fixed Ops.\nSaid: Roundtable slide 9 ('Service Department?'); Jan GM slide 58 (Sales and Service Relationship)." });

T({ eyebrow: "The CP Engine", title: "What Can You Control?", saidFrom: RT,
  said: ["ROs > Hours per RO > ELR > Mix > GP%",
         "CP labor $ per RO represents how well you execute, store and individual",
         "ELR plus hours per RO cannot be manipulated: true performance",
         "Total money in beats ELR, gross %, hours, anything else. It's what we sell."],
  now: [{ big: "279,756 ROs", label: "Customer-pay RO count, up only 456. Volume is flat.", tone: "n" },
        { big: "1.87 hrs per RO", label: "Up 0.07, ahead in all nine months. Every 0.1 hour per RO is worth $4,543,691 of labor.", tone: "g" },
        { big: "$304 per RO", label: "CP labor per RO, up $17. The $4,899,376 labor gain came from selling more per visit.", tone: "g" }],
  ask: "with car count flat, the growth is in hours per RO. Sell the next tenth of an hour on every RO.",
  tone: "g", notes: "PLACE: CP Engine section, after the section break and before CP RO Count.\nSaid: Roundtable slides 18-19 (KPIs; CP Labor $ per RO)." });

T({ eyebrow: "The CP Engine", title: "The Front Door", saidFrom: GM,
  said: ["Open and available scheduler: same-day or next-day appointments, quick online scheduling",
         "75% of visits are basic services. Saturdays are what you choose to make them.",
         "Every call is a sales opportunity. Mishandled calls cost millions.",
         "Where do calls go after hours and Sundays?",
         "No-show process; first appointment at time of sale; next appointment before the customer leaves",
         "Active retention tracking: work the Viz 'Close to Losing' and 'Unretained' lists"],
  now: [{ big: "+0.2%", label: "Customer-pay RO count. 13 of 33 stores had fewer CP ROs than last year.", tone: "n" },
        { big: "166 ROs per advisor", label: "CP & W ROs per service advisor a month, down 7. Advisors grew 1.5%; ROs fell 2.3%.", tone: "b" }],
  ask: "every customer leaves with the next appointment booked, and every call reaches someone who asks for it.",
  notes: "PLACE: CP Engine section, after the CP RO Count slides and before Hours per RO.\nSaid: Jan/May GM slides on Appointment Scheduling, Appointment Discipline, Phone Call Handling and Retention Management (Jan 2, 7, 9, 11; May 2, 4, 8, 9).\nOptional: the Flai voice AI service view (Jul 16-Aug 15 pull: 4,368 of 8,351 transfers rang out) fits here, but it is a different source and period; refresh it before use and keep it off any slide with a P&L number." });

T({ eyebrow: "The CP Engine", title: "Pricing Strategy", saidFrom: BOTH,
  said: ["Do a comprehensive market labor study",
         "Calculate your ELR on every maintenance item; target ELR by work type",
         "Know your loss leader, and what the competition is charging",
         "No direct price changes to parts or labor: use ADC-controlled discounting for every discount",
         "Benchmark: CP ELR at least 70% of door rate"],
  now: [{ big: "$162.42", label: "CP effective labor rate, up $3.56 and ahead in all nine months.", tone: "g" },
        { big: "74.7%", label: "CP ELR as a share of door rate, Jul-Sep. 25 of 33 stores at or above 70%.", tone: "g" },
        { big: "73.9% vs 76.4%", label: "Grid stores realize less of their door rate than flat-rate stores, Jul-Sep.", tone: "b" }],
  ask: "know what you collect on every op code, and run every discount through ADC so we can see it.",
  notes: "PLACE: CP Engine section, before the CP ELR slides (ahead of the door-rate options).\nSaid: Roundtable slide 15 (Pricing Strategy); Jan/May GM Road to the Sale slide (pricing lines); Jan GM slide 54 (Discounting Discipline); 70% benchmark from the Jan/May door rate slides." });

T({ eyebrow: "The CP Engine", title: "The Goldilocks Margin", saidFrom: RT,
  said: ["Gross profit % is the share of labor you keep after paying the tech",
         "Not too high, not too low: 78% to 80%",
         "Skill code costing: a hybrid of traditional and percentage",
         "Right tech, right job"],
  now: [{ big: "80.8%", label: "CP labor gross profit, down 0.1 pts. Just above the 78-80% target.", tone: "n" },
        { big: "40.7%", label: "CP parts gross profit, up 1.2 pts and ahead in all nine months.", tone: "g" }],
  ask: "hold labor margin in the 78-80% band by matching the tech to the job, not by squeezing tech pay.",
  notes: "PLACE: CP Engine section, before CP Labor Gross Profit %.\nSaid: Roundtable slide 21 (Gross Profit %). Above 80% can mean techs are underpaid on the work they do; below 78% means the job mix or the dispatch is off." });

T({ eyebrow: "Labor Mix", title: "Express Service", saidFrom: GM,
  said: ["Express service is the #1 retention driver and business builder",
         "75% of your business includes some type of maintenance work",
         "Express bays are simply MPI bays. Unknown dollars are worth 2.5 times the known.",
         "Minimize labor costs for lower-skilled work",
         "If you don't help them, someone else will"],
  now: [{ big: "54.3%", label: "Maintenance share of customer-pay labor; repair 30.5%.", tone: "n" },
        { big: "52.5%", label: "Customer pay as a share of labor sale, down 0.3 pts. Warranty labor rose 8.8%, internal 5.8%.", tone: "n" }],
  ask: "treat every express visit as an inspection: fast in, inspected, and sold the next service.",
  notes: "PLACE: Labor Mix section, after the section break and before the pay type slides.\nSaid: Jan GM slide 8 / May GM slide 7 (Express Service)." });

T({ eyebrow: "Labor Mix", title: "Tracking & Accountability", saidFrom: BOTH,
  said: ["Daily review of the Viz 'Op Code' and 'Performance Report'",
         "Give feedback quickly and often. Trends start early: don't wait until the 15th.",
         "Tracked metrics improve performance; we are responsible for the knowing",
         "Scoreboard techs and advisors: CP $ per RO, ELR, MPI completion, closing %, key op codes"],
  now: [{ big: "15.2%", label: "Customer-pay labor whose op code description can't be classified, down from 19.0%.", tone: "g" },
        { big: "4.2% to 63.5%", label: "The range across stores. Vague descriptions hide what the shop sells.", tone: "b" }],
  ask: "if it isn't described, it can't be tracked. Clean up op code descriptions so the reports show what you sell.",
  notes: "PLACE: Labor Mix section, before the work-type (maintenance / repair / unclassified) slides.\nSaid: Jan GM slide 19 / May GM slide 15 (Tracking & Accountability); Roundtable slide 16 (Performance Tracking)." });

T({ eyebrow: "Opcodes", title: "Road To The Sale", saidFrom: BOTH,
  said: ["Every advisor, every customer: timely meet and greet at the vehicle",
         "Walk-around with the customer: hood, tires and wipers",
         "Review history and previously declined items",
         "Time and mileage recommendations from an accurate, available menu",
         "MPI mandate: every car, every time",
         "Consistency is key: trained, documented and actively managed"],
  now: [{ big: "Up in all four", label: "Alignment, balance, brake fluid and fuel service each rose 1.2-1.4 pts and beat 2025 every month.", tone: "g" },
        { big: "18,493", label: "Balances short of the 15% benchmark. The group is at 8.3%.", tone: "b" },
        { big: "4,821", label: "Alignments short of 15%; the group is at 13.2%. Brake fluid is 0.7 pts from its 10%.", tone: "b" }],
  ask: "every car gets the walk-around and the menu. The benchmarks are where that process lands.",
  notes: "PLACE: Opcodes section, after the section break and before Alignment %.\nSaid: Roundtable slides 11, 12, 14 (Road to the Sale, Maintenance Menus, MPI Mandate); Jan GM slide 15 / May GM slide 11 (Road to the Sale). Fuel service is 5.8% vs 7.5% (4,731 short)." });

T({ eyebrow: "Warranty & Internal", title: "Warranty Compliance", saidFrom: JAN,
  said: ["Higher labor rates and tariffs are driving up OEM scrutiny",
         "Six audits cost $400,000",
         "Primary causes: written documentation, manager signatures on add-ons and in-stock units, out mileage, punch times, manufacturer authorization",
         "The general manager is responsible for warranty policy being followed",
         "Spot-check repair orders and self-audit"],
  now: [{ big: "$197.17", label: "Warranty ELR, up $16.93; 30 of 33 stores raised it.", tone: "g" },
        { big: "−8,873", label: "Warranty ROs, down to 126,891. Volume follows OEM campaigns and recalls.", tone: "n" }],
  ask: "a higher warranty rate brings more scrutiny. Self-audit before the factory does.",
  notes: "PLACE: Warranty & Internal section, after the section break and before Warranty RO Count.\nSaid: Jan GM slides 13-14 (audit chargebacks; Warranty Compliance). The Jan slide named six stores' chargebacks; kept to the $400,000 total here, in line with no store call-outs from the podium." });

T({ eyebrow: "Fixed Ops Productivity", title: "People", saidFrom: BOTH,
  said: ["Our most valuable asset: this is what separates us",
         "Invest in training and a defined career path. Be the best place to work.",
         "Grow your own: apprentice programs",
         "What does being short-handed cost? Limited scheduling, lost opportunities, missed MPIs",
         "Clear pay plans, job descriptions and expectations"],
  now: [{ big: "−76", label: "Fixed ops staff, Sep to Sep, to 1,540 (−4.7%): technicians −17, advisors −7, managers −3.", tone: "n" },
        { big: "$13,617", label: "Fixed gross per fixed employee a month, up $1,207. Gross grew 4.1% with 5.1% fewer people.", tone: "g" },
        { big: "−2.3%", label: "CP & W ROs year to date. Fewer people and fewer cars: the next gain has to come from car count.", tone: "b" }],
  ask: "keep every seat filled with someone who adds hours. Being short-handed costs appointments, MPIs and sold hours.",
  notes: "PLACE: Productivity section, after the section break and before Fixed Gross per Fixed Employee.\nSaid: Roundtable slide 6 (People); Jan GM slide 57 (Technician Pipeline).\nNumbers rebuilt 2026-10-20 on ADP headcount (Employee Master extract 10/19/2026); the Qlik headcount history undercounted 2025." });

T({ eyebrow: "Fixed Ops Productivity", title: "Pay Plans Are Work Instructions", saidFrom: GM,
  said: ["Clear and easily understood",
         "Based on customer-pay performance: CP labor or gross per RO, maintenance penetration, total sales or gross",
         "Drive the results you are looking for",
         "Base versus bonus: a 50/50 ratio",
         "Every system is perfectly designed to get the results it gets"],
  now: [{ big: "$304", label: "CP labor per RO, the number the plans should pay on. Up $17.", tone: "g" },
        { big: "166", label: "CP & W ROs per advisor a month, down 7. Advisors grew 1.5%; ROs fell 2.3%.", tone: "b" }],
  ask: "check that every advisor and tech pay plan pays for the numbers on these slides.",
  notes: "PLACE: Productivity section, before CP & W ROs per Service Advisor.\nSaid: Jan GM slide 16 / May GM slide 12 (Pay Plans aka Work Instructions)." });

T({ eyebrow: "Technician Proficiency", title: "Technicians & Stalls", saidFrom: BOTH,
  said: ["1:1 tech to stall unless you run on-demand appointments",
         "Tech level in line with job skill: right tech, right job",
         "5/6 of staff on schedule at all times, Saturdays included",
         "90%+ MPI completion; track and report technician efficiency",
         "Use scoreboarding and Teams to show techs where they stand"],
  now: [{ big: "94.7%", label: "Group technician proficiency (sold hours / paid hours). 20 of 33 stores at or above 85%.", tone: "g" },
        { big: "262 of 583", label: "Counted technicians under 85%. Every store has at least one.", tone: "b" },
        { big: "$1.65M a month", label: "Gross missed if each of them reached 85%: no extra hours, techs or bays.", tone: "b" }],
  ask: "show every tech their own proficiency every week. The best techs in each store already prove 85% is possible.",
  notes: "PLACE: Technician Proficiency section, after the section break and before the store proficiency ranker.\nSaid: Roundtable slide 22 (Technicians & Stalls); Jan GM slide 52 / May GM slide 16 (MPI Completion and Tech Efficiency).\nNow: proficiency slides (ADP paid hours + Qlik sold hours, Dec 25-Oct 9, with the foreman 30%, tech 10%, role-change and keying-error rules)." });

T({ eyebrow: "Fixed Ops", title: "Before You Leave", saidFrom: RT,
  said: ["What did you learn today?",
         "Give us one or two things you are going to act on and implement",
         "Do you want to do this again?",
         "Reach this room in FMGSVCMGR"],
  nowLabel: "Where the money is",
  now: [{ big: "0.1 hr per RO", label: "Worth $4,543,691 of labor at today's ELR.", tone: "g" },
        { big: "18,493 balances", label: "The largest benchmark gap, at 8.3% against 15%.", tone: "b" },
        { big: "$1.65M a month", label: "Technician proficiency, lifting each tech under 85% to 85%.", tone: "b" }],
  ask: "name your one or two commitments out loud, with a number attached, before you leave the room.",
  tone: "g", notes: "PLACE: the close, after the last Productivity / Proficiency slide.\nSaid: Roundtable slide 24 (Misc & Closing). Commitments tracking was dropped for this meeting, so this asks for commitments without a scorecard." });

await pptx.writeFile({ fileName: "Talking Points - Day 2.pptx" }); console.log("tp built", pptx._slides.length);
