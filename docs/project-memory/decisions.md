# Decisions

- Keep the stable project/plugin identity `native-powerpoint-doc-editor`; the manifest ID is the release-stable API identifier.
- Use npm scripts and esbuild for build, typecheck, lint, and verification. Source TypeScript remains under `src/`; bundled `main.js` is a generated release artifact.
- Resolve the editor theme at the plugin level and pass the resolved value to DOCX/PPTX consumers. Theme colors use `--npde-*` tokens; component rules do not introduce hardcoded color literals.
- Keep DOCX and PowerPoint handling independently disableable through plugin settings, so another plugin can own either extension.
- Keep normal operation local/offline. Network access is limited to explicit external-link use or DOCX exports whose images reference remote URLs; there is no telemetry or self-updating code.
