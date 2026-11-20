# Inventory & Order Tracking Automation (Excel VBA)

A macro-enabled Excel workbook that runs inventory and order tracking end to end: product, supplier and order entry through validated VBA forms, formula-driven stock and reorder calculations, and PivotTable dashboards for inventory value, reorder alerts and order status.

## What it does

- **Data entry with guardrails.** VBA UserForms for product, supplier and order entry, with required-field and numeric validation, sequential ID generation, and save/clear actions.
- **Stock and reorder logic.** INDEX/MATCH, SUMIFS and IF formulas calculate delivered quantities, remaining stock, inventory value, low-stock alerts and replenishment quantities against per-product reorder thresholds.
- **Dashboards.** Two PivotTable-based dashboards with charts and slicers covering inventory value, stock availability, reorder alerts, supplier lead times and delivered vs pending orders.

## Data model

Structured Excel tables link **50 products**, **10 suppliers** and **250 order records** across 14 sheets, so every entry flows through to stock positions, reorder reports and dashboards without manual rework.

## Tools

Excel (structured tables, PivotTables, PivotCharts, slicers, conditional formatting), VBA (UserForms, event handlers, report macros)

## Why it matters

This mirrors the day-to-day of procurement and inventory analysis: keeping master data clean, knowing what to reorder and when, and giving planners numbers they can trust.
