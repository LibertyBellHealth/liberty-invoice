# Staging environment

This `staging` branch powers the **permanent staging preview** for the Home Care CRM.

- It keeps one long-lived pull request open against `main`, which gives Azure Static
  Web Apps a **fixed preview URL** that does not change between changes.
- That fixed URL is registered once as an allowed sign-in address (Azure AD) and an
  allowed CORS origin on the backend, so previews work without per-change setup.

**Do not merge this PR into `main`.** Closing/merging it would tear down the staging
environment and change the URL. Features ship to production by merging their own
branch into `main`, not by merging `staging`.

To preview a change: it gets merged into `staging` (redeploys the same fixed URL).
Once approved, the change is merged into `main` to go live.
