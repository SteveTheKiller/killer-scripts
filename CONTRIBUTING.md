# Contributing to killer-scripts

I handle most development myself. Bug reports and feature requests help me decide what to work on, but opening an issue does not assign the implementation to its author.

## Before you start

Search the existing issues first. Put feature requests in an issue and explain the problem, the intended workflow, and why the current behavior falls short. Wait for me to approve the scope and invite implementation before starting feature work or opening a feature PR.

The same approval is required for architectural changes, new dependencies, substantial UI rewrites, changes to data preservation or compatibility, and work intended for a future milestone. A draft PR is still a PR; please discuss the proposal first.

You do not need advance permission for:

- A focused bug fix that preserves the intended behavior
- A regression test
- Translation work
- A documentation correction
- A small theme, layout, or accessibility fix that preserves the established design

If a small fix grows into a redesign, pause and discuss the larger change in an issue. Unrequested features, speculative rewrites, and large chains of dependent PRs may be declined without a full code review. Approval to work on an issue does not guarantee that a PR will be merged.

## Keep changes focused

Address one coherent problem per PR. Keep unrelated formatting, cleanup, file moves, dependency updates, and refactoring separate.

Preserve behavior outside the approved scope. Reuse the existing architecture, controls, styles, spacing, and resources. Explain the specific problem before changing an established design or a boundary between components.

Do not bump versions, change release dates, prepare packages, or rewrite release notes unless the contribution is specifically approved release work. Ask which branch to target before starting an approved feature.

## Testing

Follow the repository's build and test instructions. Add a focused regression test for a reproducible bug when there is a useful test boundary.

Test the actual workflow affected by the change. A successful build alone does not prove that a UI, save path, installer, or hardware interaction works. For visible changes, check keyboard and mouse use, supported themes, scaling, and narrow windows where applicable. Include before-and-after images when they help explain the change.

State which checks passed and which paths you could not test. Use fabricated or safely redacted fixtures, and keep private data and credentials out of issues, logs, and test files.

Translation-only and documentation-only changes do not need a full application build unless they affect generated files or runtime behavior. Preserve resource keys and format placeholders, and follow the repository's translation guide when one exists.

## Dependencies and data safety

Discuss new or replacement dependencies before adding them. Explain why they are needed, their license, maintenance status, platform support, size, native requirements, and security or update impact.

Do not suppress vulnerability warnings or weaken validation to make a sample pass. Preserve existing files and user data when an operation fails, and explain any compatibility or data-loss risk.

## Pull request descriptions

Include:

- The problem and the related issue
- What changed and why
- How you verified it
- Images or small sample files when relevant
- Known limitations and untested behavior
- The approval for feature work or other substantial changes

Avoid automatic issue-closing keywords in PR descriptions and commits. Bug reports stay open for reporter confirmation when practical.

Be prepared to revise an approved contribution after review. Contributions must be compatible with the repository's existing license; these guidelines do not change that license.

## Project-specific guidance

Describe the machine changes a script makes and how failures are handled. Use a disposable environment for scripts that change system settings or remove software or data.

