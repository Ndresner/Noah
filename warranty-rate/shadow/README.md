# Shadow run – in-house rate vs. DealerUplift, live cycles

Started Nov 24, 2026. For each DealerUplift (Armatus) project, we calculate our own statutory rate from Qlik (180-day look-back ending at the Armatus project start date, 100 sequential qualifying customer-pay ROs, state rules) **before** their preliminary results arrive, then log their numbers when they do.

- `shadow_runs.csv` – one row per store/cycle: our best / median / latest window, DealerUplift prelim and final, break-even tolerance, manufacturer outcome. Aggregate rates only.
- `shadow_config.json` – store labor-type config and cycle dates used by `scripts/history_vs_dealeruplift.py`.
- Break-even tolerance per hour = (10% of DealerUplift's annual lift + $2,250) / annual warranty hours (hours = their quoted annual value / increase). In-house wins if our rate ≥ their rate − tolerance.
- A scheduled check (weekdays, twice a day) reads the mailbox for new DealerUplift preliminary/final emails and reports the comparison. It never replies to or sends email.

Upcoming Armatus starts (from Noah's 11/13 list): Kia (LV) 11/19, Subaru LV 11/20, Jaguar LV & Jaguar Henderson 11/25, Honda Spokane 12/02, Findlay Motor Company (GM) 12/07, Hyundai Prescott 12/10, VW St George 12/14.
