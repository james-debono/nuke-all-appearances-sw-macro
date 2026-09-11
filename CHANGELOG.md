# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

## 0.1.0 — 2026-09-11

- First release. Clears face, feature, body, part and component appearances from
  every display state and every configuration, and from every part document an
  assembly references.
- Referenced parts are modified in memory and marked unsaved. Nothing is written
  to disk, so closing without saving reverses the whole run.
- Suppressed and lightweight components are resolved first, since an appearance
  cannot be cleared from a component that is not loaded. They stay resolved.
- Asks for confirmation before running, and states in the prompt that referenced
  part files will be changed.
