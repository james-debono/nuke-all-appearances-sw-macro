# Nuke All Appearances — development notes

The API findings behind this macro. `README.md` covers what it does and how to
use it.

## What it clears

Every appearance, on both of the axes SOLIDWORKS scopes them by:

| Axis | Reach |
|---|---|
| Level | face, feature, body, part, component |
| Display state | all of them |
| Configuration | all of them |
| Document | the open one, plus every part an assembly references |

## The dead end that became the mechanism

**Read this before "fixing" anything here.**

[Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro)
and [Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro)
both record `RemoveMaterialProperty` as a dead end from their 0.1.0:

> Scoped by *configuration*, and several display states live under one
> configuration, so it cleared them all.

That was a defect in those macros. **Here it is the entire point.** Both of them
then moved to `GetRenderMaterials2` + `AddDisplayStateSpecificRenderMaterial`
precisely because those can be pointed at one display state — which is the
behaviour this macro exists to avoid.

So the rejected approach is the correct one in this repository, and
`swAllConfiguration` widens it further to cover every configuration as well.
Anyone porting a fix between these three macros needs to know that this one is
deliberately upside-down relative to its siblings.

## Two sweeps, and why both

Appearances are scoped on two independent axes, and neither sweep provably covers
both. A macro called Nuke earns the redundancy.

### Sweep A — render materials, across all display states

```vba
vAppearances = swModel.Extension.GetRenderMaterials2( _
    swDisplayStateOpts_e.swAllDisplayState, Empty)
' then, per appearance:
swAppearance.RemoveAllEntities
```

An appearance is a render material plus the list of entities it is attached to.
An appearance left with no entities is gone.

**Nothing is ever written back**, because nothing is ever kept. That is what makes
this cheap: Remove All Appearances 0.3.0 spent **161 seconds re-attaching 5,359
faces**, and that failure mode cannot occur here — there is no re-attach step at
all.

This is also the **only** sweep that reaches component appearances, since an
assembly document has no bodies, faces or features of its own.

### Sweep B — `RemoveMaterialProperty`, across all configurations

CodeStack's `remove-color` pattern, at four levels:

```vba
swModel.Extension.RemoveMaterialProperty swAllConfiguration, Empty
swBody.RemoveMaterialProperty           swAllConfiguration, Empty
swFace.RemoveMaterialProperty2          swAllConfiguration, Empty
swFeat.RemoveMaterialProperty2          swAllConfiguration, Empty
```

Catches anything scoped by configuration that sweep A's display-state view does
not see. Note the `2` suffix on the face and feature calls and its absence on the
body and document ones — that is the API, not a typo.

## The per-face sweep is the expensive one

`NUKE_FACES_INDIVIDUALLY` controls one COM call per face, and a large weldment has
thousands. **Sweep A already clears face appearances wholesale**, one call per
appearance rather than per face, so this constant is redundancy rather than the
main mechanism.

If a big assembly crawls, this is the first thing to turn off. Turn
`SHOW_DIAGNOSTICS` on first and read the timing breakdown rather than guessing —
the dialog reports faces swept alongside the nuke time.

## Nothing is saved, and that is the whole safety model

`SetSaveFlag` marks each modified document dirty so SOLIDWORKS prompts for it on
close. Nothing is written to disk by the macro.

This is the only way back. Appearance changes leave **no undo record**, so once a
part is saved the colours are gone. Closing without saving reverses the entire
run — including the component resolves, since suppression state is stored in the
assembly and is equally unsaved.

The two decisions reinforce each other: because nothing is written, resolving every
component is as reversible as the appearance removal itself.

## Resolving components — verify the result, not the return

```vba
swComp.SetSuppression2 swComponentSuppressionState_e.swComponentResolved
' then re-read it:
If swComp.GetSuppression() = swComponentSuppressed Then  ' it failed
```

The call's return enum is deliberately not trusted. A resolve that silently failed
would leave a part quietly uncleared while the summary still claimed success,
which is the worst outcome available — the user would save, believing it worked.

Lightweight components go in one call, `ResolveAllLightWeightComponents True`,
before the individual pass.

## De-duplicating documents

Documents are keyed by lower-cased `GetPathName()`. A part placed fifty times is
one document and is nuked once. Doing it fifty times would not be *wrong*, just
fifty times slower — but on a large assembly that is the difference between
seconds and minutes.

The open document is seeded into the dictionary before the walk, so an assembly
that somehow references itself cannot be processed twice.

## Propagate visual properties — read this before debugging appearances

Pattern and mirror features have a **"Propagate visual properties"** option, on by
default. When on, regenerating the feature stamps the seed body's colour onto the
derived bodies' **faces**. A face appearance beats a body appearance, so colour can
reappear after a rebuild even though removal succeeded.

**Check that checkbox before suspecting the macro.** This trap is recorded in both
sibling repositories too, because a forker of one of them must not have to find it
twice.

## `EditRebuild3` is not optional

Without it the viewport keeps drawing the old colours until the display state is
switched away and back.

## Known limitations

- **No undo.** Closing without saving is the only way back.
- **Read-only and PDM-controlled parts** are modified in memory but cannot be
  saved. Nothing warns the user until they try — the macro has no way to know in
  advance without probing every file.
- **Drawings are not supported**, and are rejected up front.
- Feature-level failures are ignored rather than counted: not every feature type
  carries an appearance, so the failures are expected and reporting them would be
  noise.

## Verification status

Stated precisely, because at 0.1.0 it is uneven.

**Verified without running anything:** the API surface. Every call used here
appears either in a working sibling macro in this folder
(`GetRenderMaterials2`, `RemoveAllEntities`, `GetEntities`, `EditRebuild3`,
`EnableGraphicsUpdate`) or in CodeStack's published examples
(`RemoveMaterialProperty`, `RemoveMaterialProperty2`, `GetBodies2`, `GetFaces`,
`GetFeatures`, `GetComponents(False)`, `ResolveAllLightWeightComponents`,
`GetModelDoc2`, `GetPathName`, `SetSaveFlag`, `GetSuppression`).

**Assumed, and the most likely thing to need a fix:** `SetSuppression2` with
`swComponentSuppressionState_e.swComponentResolved`. It is used here from the
documented API rather than from a working example in either source, which is why
the code re-reads `GetSuppression()` afterwards instead of trusting it.

**Not yet confirmed in SOLIDWORKS:** everything else. Nothing that touches the
SOLIDWORKS API can be verified any other way, and this section will say so until
it has been run.

## There is no build step

A `.swp` is a binary VBA project. Editing the `.vba` in `src\` changes nothing
that runs until the source is pasted into the SOLIDWORKS VBA editor and saved.
Treat the `.vba` as the source of truth and re-paste after every change.

Do not patch a `.swp` directly: it stores compiled p-code ahead of the source
text and VBA runs the p-code, so a patched file shows new code in the editor
while still running the old.
