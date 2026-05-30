# =============================================================================
# setup_site.R
#
# Run ONCE locally to bake your chosen password into the dashboard.
# The password is hashed with SHA-256 (client-side, same algo in the JS).
# No plaintext password is ever stored in the repo.
#
# Usage:
#   Rscript setup_site.R
# Then commit & push docs/index.html.
# =============================================================================

# ── Choose your password ──────────────────────────────────────────────────────
# Change this to whatever you want your group's shared password to be.
SITE_PASSWORD <- "mlb2026"

# ── Hash it (must match the JS: SHA-256 of password + 'manubets_salt_2026') ──
if (!requireNamespace("digest", quietly = TRUE)) install.packages("digest")

hash_password <- function(pw) {
  digest::digest(paste0(pw, "manubets_salt_2026"), algo = "sha256", serialize = FALSE)
}

pw_hash <- hash_password(SITE_PASSWORD)
message("Password hash: ", pw_hash)

# ── Patch into docs/index.html ────────────────────────────────────────────────
html_path <- "docs/index.html"
if (!file.exists(html_path)) stop("docs/index.html not found. Run from repo root.")

html <- readLines(html_path, warn = FALSE)
html <- gsub("%%PASSWORD_HASH%%", pw_hash, html, fixed = TRUE)
writeLines(html, html_path)
message("✓ Password hash written to docs/index.html")

# ── Also store in GitHub Actions secret reminder ──────────────────────────────
message("\n── GitHub Secrets to add (repo Settings → Secrets → Actions): ──")
message("  SHEET_ID                = 1F3hHYptA-lvD3o8a5P3y46yN_uLn5BlqwSjDA4yeUC0")
message("  GS_SERVICE_ACCOUNT_JSON = (contents of your service account .json file)")
message("\n── Tell your group: ──")
message("  URL:      https://manulbets-web.github.io/mlb-model")
message("  Password: ", SITE_PASSWORD)
message("\n── Next steps: ──")
message("  1. Commit and push docs/index.html")
message("  2. GitHub repo Settings → Pages → Source: Deploy from branch 'main', folder '/docs'")
message("  3. Add GitHub Secrets (see above)")
message("  4. Run the Actions workflow manually once to generate the first data.json")
