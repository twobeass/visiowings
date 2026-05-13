<!--
Thanks for contributing to visiowings!

PR titles must follow Conventional Commits:
  feat:      new feature       (-> minor version bump)
  fix:       bug fix            (-> patch version bump)
  feat!:     breaking change    (-> major version bump after 1.0)
  docs:      documentation
  test:      tests only
  refactor:  no behavior change
  perf:      performance
  build:     build / packaging
  ci:        CI / pipeline
  chore:     misc maintenance

Example: feat(cli): add `visiowings init` wizard
-->

## Summary

<!-- 1-3 bullet points describing what this PR does and why. -->

## Type of change

- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation
- [ ] Refactor / internal cleanup
- [ ] Tests only
- [ ] CI / build

## How was this tested?

<!--
- [ ] `pytest` runs green
- [ ] Manually verified on Windows + Visio
- [ ] N/A
-->

## Checklist

- [ ] PR title follows Conventional Commits
- [ ] Tests added / updated
- [ ] Docstrings / docs updated where applicable
- [ ] No new `print()` calls (use `logging`)
- [ ] No bare `except:` clauses
- [ ] Pre-commit hooks pass locally (`pre-commit run --all-files`)
