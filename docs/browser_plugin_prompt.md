# Prompt: Add browser-based signing UX with CryptoPro plugin integration

Use this prompt as guidance for implementing an optional browser-based signing flow that leverages the installed CryptoPro CSP browser plug-in instead of the current in-app PKCS#11 dialog.

## Context
- App: desktop Python/PySide6 tool for visualising and stamping electronic signatures in PDFs.
- Current signing: in-app dialog with two tabs — certificates from files or PKCS#11 token (user provides library path, PIN, optional slot/label).
- Goal: add a streamlined UX similar to "choose certificate in browser → enter token password → sign" using the CryptoPro browser plug-in already available in the user's browser.

## UX requirements
1. Provide a new signing option (e.g., "Подписать через браузер (CryptoPro)") alongside the existing dialog.
2. When chosen, open the user’s default browser pointing to a local page hosted by the app (loopback HTTP). The page should:
   - Prompt the user to pick a certificate via the CryptoPro plug-in UI (no manual provider selection).
   - Request PIN/password when the plug-in asks and apply the signature.
   - Return the raw detached PKCS#7 signature for the PDF hash.
3. After the browser confirms success, the desktop app continues the existing flow: verifying the signature, showing certificate info, and allowing stamp placement/saving.
4. Preserve the current in-app signing modes; browser signing is additive and optional.

## Technical guidelines
- Spin up a temporary local HTTP endpoint (127.0.0.1, random free port) only during the browser-based signing session; shut it down afterward.
- Implement a minimal HTML/JS page that calls the CryptoPro plug-in: invoke `cadesplugin` to access certificates, prompt selection, compute the CMS signature for the provided hash (aligned with the PDF being signed), and POST the result back to the app endpoint.
- Ensure CSRF/loopback safety: restrict host to localhost, use a one-time nonce in both the URL and payload, and reject mismatched requests.
- Handle errors clearly: surface plug-in errors, cancelled selection, missing plug-in, or unsupported browser with user-friendly messages in the desktop UI.
- Keep platform compatibility in mind (Windows primary target for CryptoPro plug-in; fail gracefully on other OSs).

## Acceptance criteria
- New UI affordance to launch browser signing and status messaging for success/failure.
- Local server lifecycle is automatic; no manual port configuration.
- Signature returned from the browser path is validated and piped into existing PDF signing/verification logic.
- Clear fallbacks: if browser signing fails, users can retry or switch back to file/PKCS#11 modes without restarting the app.

Use this prompt as the starting point when implementing the feature or guiding another agent to do so.
