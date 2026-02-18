# Roadmap

## Completed
- [x] Task pane UI with date preview and email preview
- [x] Auto-calculated booking date (today + 28 days)
- [x] Pre-filled compose window via `displayNewMessageForm()`
- [x] Bullet-point booking details (name, residences, room, date, duration, mobile)
- [x] Mailto fallback for browser testing
- [x] HTTPS dev server setup with office-addin-dev-certs

## Planned

### Recurring Schedule
- [ ] Outlook rule or Power Automate flow to trigger every Thursday at 8am
- [ ] Auto-open the add-in on schedule so Steve only needs to review and send

### Editable Fields
- [ ] Allow editing booking details (date, duration, room type) in the task pane before composing
- [ ] Date picker for manual override of the 28-day default

### Multiple Bookings
- [ ] Support booking multiple dates in a single email
- [ ] Week-at-a-glance view showing upcoming bookings

### Deployment
- [ ] Host on Azure Static Web Apps or GitHub Pages (eliminate local server requirement)
- [ ] Centralized deployment via Microsoft 365 admin center
- [ ] Replace sideloading with org-wide add-in distribution

### UX Improvements
- [ ] Confirmation toast after compose window opens
- [ ] History log of previously sent booking requests
- [ ] Dark mode support matching Outlook theme

### Unified Manifest
- [ ] Migrate from XML manifest to the Teams JSON manifest format for future compatibility
