# Email Reviewer design system

## Product context
- Product: offline email evidence review utility
- Audience: legal, compliance, IT, and investigation staff
- Platform: Windows 10/11 desktop

## Stack
- .NET 8 WinForms
- Native controls and custom-painted controls only; no web-style framework assumptions

## Visual direction
- Direction: Evidence Desk
- Keywords: restrained, trustworthy, dense, highly scannable
- Shape: 8 px action-button and input corners; square data grids
- Elevation: flat surfaces separated by color and two-pixel control borders

## Semantic tokens
- Primary: navy for application identity and grid headings
- Accent: blue for primary actions and selection
- Canvas: cool light gray
- Surface: white
- Subtle surface: pale blue-gray
- Text: charcoal
- Muted text: slate
- Typography: Segoe UI Variable Display for display roles; Segoe UI Variable Text with Segoe UI fallbacks
- Spacing: 4/8-based rhythm; 16 px section padding
- Motion: immediate state changes; hover and pressed color feedback only

## Layout rules
- Preserve the three-pane workflow: folders, sortable message grid, reading pane
- Preserve Date, From, To, Subject, and Preview headings
- Keep filters above all three panes
- Keep Messages and its live filtered Total visually distinct and aligned
- Minimum application size: 1200 × 700

## Components
- Rounded action buttons: full rectangular paint surfaces behind anti-aliased shapes prevent clipped edges and corner artifacts
- Inputs: persistent labels, rounded two-pixel borders, consistent 36 px minimum height; date-filter checkboxes sit outside the bordered date field
- Message grid: SQL-backed sorting, high-contrast headings, stable column names
- Reading pane: display-type subject, compact metadata, body-text typography optimized for long reading

## Accessibility
- Windows High Contrast uses system colors and native button painting
- Every primary action remains keyboard reachable
- Do not remove focus indicators
- Do not communicate status through color alone

## Anti-patterns
- Decorative gradients, shadows, or animation
- Rounded data grids or card-heavy layouts
- Removing sortable message-column headings
- Replacing persistent labels with placeholders

## Review checklist
- Test at 1200 × 700 and high DPI
- Verify keyboard focus and High Contrast behavior
- Confirm all message columns, filters, paging, and reading behavior remain unchanged
- Run the complete Email Review Viewer test suite before publishing
