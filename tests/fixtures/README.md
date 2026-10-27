# Retired metadata rejection fixtures

These synthetic Office files were captured from snapshot `b97b7eaf`. Their bytes remain unchanged, but the editor no longer interprets their old app-owned metadata as editable component objects. They are negative inputs for foreign-object protection tests, not a supported compatibility format.

- `retired-namespace-smartart.xlsx`: an old DiagramML graph and opaque package content.
- `retired-namespace-model3d.xlsx`: an old 3D anchor with its original closed GLB.

Arbitrary user text stays unchanged. New objects use the current component namespace. Historical revision/hash fixtures and dual-dialect readers were retired with the clean-start baseline.
