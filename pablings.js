#!/usr/bin/env node
"use strict";

const readline = require("readline");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function ask(q) {
  return new Promise((resolve) => rl.question(q, (ans) => resolve(ans.trim())));
}

function toNum(label, s) {
  const n = Number(String(s).replace(/,/g, ""));
  if (!Number.isFinite(n)) throw new Error(`Invalid number for ${label}: ${s}`);
  return n;
}

function peso(n) {
  return "₱" + Math.round(n).toLocaleString("en-PH");
}

(async () => {
  try {
    console.log("\n=== Barbershop Franchise ROI Quick Calculator ===\n");
    console.log("Tip: you can type numbers with commas, e.g., 150,000\n");

    const capex = toNum("CAPEX", await ask("1) Total initial investment (CAPEX) ₱: "));
    const monthlySales = toNum("Monthly Sales", await ask("2) Expected average monthly SALES ₱: "));

    // Cost of goods / supplies (wax, towels laundry, products, etc.)
    const cogsPct = toNum("COGS %", await ask("3) Supplies/COGS (% of sales) e.g. 8: "));
    const rent = toNum("Rent", await ask("4) Monthly RENT ₱: "));
    const payroll = toNum("Payroll", await ask("5) Monthly PAYROLL total ₱ (wages+commissions+benefits): "));

    const utilities = toNum("Utilities", await ask("6) Utilities (power/water/net) ₱: "));
    const otherOpex = toNum("Other OPEX", await ask("7) Other monthly OPEX ₱ (repairs, laundry, etc.): "));

    const royaltyPct = toNum("Royalty %", await ask("8) Royalty fee (% of sales) e.g. 5 (type 0 if none): "));
    const marketingPct = toNum("Marketing %", await ask("9) Marketing fund (% of sales) e.g. 2 (type 0 if none): "));

    const cogs = (cogsPct / 100) * monthlySales;
    const royalty = (royaltyPct / 100) * monthlySales;
    const marketing = (marketingPct / 100) * monthlySales;

    const totalOpex = rent + payroll + utilities + otherOpex + cogs + royalty + marketing;
    const net = monthlySales - totalOpex;

    // Payback (months) only meaningful if net > 0
    const paybackMonths = net > 0 ? capex / net : Infinity;
    const roiAnnual = net > 0 ? (net * 12) / capex : -Infinity;

    // Break-even monthly sales (solve for sales where net=0):
    // net = sales - (fixed + pct*sales) => net = sales*(1-pct) - fixed
    const pctCosts = (cogsPct + royaltyPct + marketingPct) / 100;
    const fixedCosts = rent + payroll + utilities + otherOpex;
    const breakEvenSales = (1 - pctCosts) > 0 ? fixedCosts / (1 - pctCosts) : Infinity;

    console.log("\n=== Results ===");
    console.log("Monthly sales:         ", peso(monthlySales));
    console.log("Monthly total costs:   ", peso(totalOpex));
    console.log("Monthly net (estimate):", peso(net));

    console.log("\nBreak-even sales/month:", Number.isFinite(breakEvenSales) ? peso(breakEvenSales) : "N/A (percent costs >= 100%)");

    if (Number.isFinite(paybackMonths)) {
      console.log("Payback period:        ", paybackMonths.toFixed(1), "months");
      console.log("Simple annual ROI:     ", (roiAnnual * 100).toFixed(1) + "%");
    } else {
      console.log("Payback period:         N/A (net is <= 0)");
      console.log("Simple annual ROI:      N/A (net is <= 0)");
    }

    console.log("\nNotes:");
    console.log("- This is a rough model. Taxes, depreciation, loan interest, and one-time repairs are not included.");
    console.log("- Use this during the meeting: plug in their numbers and see if it still makes sense.\n");

    rl.close();
  } catch (err) {
    console.error("\nError:", err.message);
    rl.close();
    process.exit(1);
  }
})();