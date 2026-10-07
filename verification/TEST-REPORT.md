# Verification and limits

## Completed

- 30 Node.js unit tests passed. These cover the existing note model, simulated Office custom-properties adapter, and six hosted manifest-builder cases.
- 7 offline Chromium DOM checks passed. These cover setup page state, generated XML download bytes and filename, resource-path mapping, a narrow layout, missing-file handling, and rejection of localhost.
- All website JavaScript files passed Node's syntax checks.
- The generated XML was parsed with lxml and checked for its original add-in ID, version 1.0.1.0, ReadItem permission, nested VersionOverrides, SupportsPinning, and URLs pointing to existing packaged website files.
- The setup-page screenshot was visually inspected.

## Not established

- No live GitHub site was created or published.
- No Outlook/Exchange mailbox was connected, and no actual note persistence was tested against Microsoft 365.
- The PowerShell installer was not run. Cmdlet syntax and usage were reviewed against Microsoft documentation; this environment does not include PowerShell.
- Microsoft's office-addin-manifest validator was not available. An npm registry check failed due to DNS/network unavailability. XML parsing and local structure tests are NOT Microsoft schema validation or Marketplace certification.
- Browser navigation is blocked by this environment's administrator policy. Offline DOM tests used about:blank, in-memory production JavaScript with the location/import wiring adjusted only in the test harness, mocked fetch responses, and an intercepted download action. They do not verify real HTTPS requests, GitHub deployment, or actual browser downloading.

The production site files contain none of the test mocks. Only publish the contents of site/. The example manifest under verification/ uses a fictitious host for tests and must not be installed.

Run the live pilot checklist in README.md after publication and mailbox installation. If Message Notes is still absent, inspect the exact mailbox's registration and Enabled state, then compare a normal received message in Outlook on the web and new Outlook. Hosting alone is not a confirmed fix for the missing button.
