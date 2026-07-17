# SDD Progress — Plan 1: foundation + read
Task 1: complete (commits c1d82dd..fa4d1b9, review clean)
  Minor (deferred to final review): bun-types declared in tsconfig but not a devDependency (matters only under tsc --noEmit); .superpowers/sdd artifacts got committed rather than gitignored.
Task 2: complete (commits fa4d1b9..8ed217c, review clean)
  Minor (deferred): leading-zero IDs normalize silently (p:04); ledger notes ride in code commits.
Task 3: complete (commits 8ed217c..5e13cf3 incl. hint fix, re-review clean)
  Decision: plan-mandated stderr-as-hint finding fixed without user escalation — Global Constraints ("hints state recovery path") govern over example code by the plan's own terms.
  Minor (deferred): ensureVenv cold-cache race; version check after venv creation leaves fingerprint-less dir on PYTHON_TOO_OLD; PYTHON_TOO_OLD reuses PYTHON_MISSING message text; long throw lines may rewrap under Prettier.
Task 4: complete (commits 5e13cf3..853ea01 incl. hint+pycache fix, re-review clean)
  Minor (deferred): _worker.py PYTHON_EXCEPTION hint generic; timedOut flag race at exact timeoutMs boundary; long template-literal throw lines.
Task 5: complete (commits 853ea01..02eb539, review clean)
  Minor (deferred): fixtures.ts pipes stdout without draining (fine while gen script prints one line).
