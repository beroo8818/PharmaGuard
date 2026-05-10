PharmaGuard fixed build

What was fixed:
1) Analytics Import now automatically creates/updates monthly_outputs/monthly_reports_latest.xlsx.
2) Executive + Storekeeper Dashboard reads the latest monthly report after Refresh.
3) Reports Hub sends one ZIP attachment to n8n using binary field name report_file.
4) n8n config accepts n8n_email_config.json values.
5) Added reports/import_sync_log.txt for import/dashboard sync logs.
6) Added START_PHARMAGUARD_SILENT.vbs to open without the black console window.

Beginner usage:
- Run START_PHARMAGUARD_SILENT.vbs
- Open Analytics and import Excel
- Open Executive Dashboard and press Refresh
- Open Reports Hub -> Create Email Summary Package -> Send Package to n8n
