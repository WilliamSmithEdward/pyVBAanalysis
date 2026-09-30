# Test fixtures

Binary containers the tests cannot author themselves. Everything else is built
in the test that needs it, with pyOpenVBA.

- `PowerPointFixture.ppt`: a legacy PowerPoint presentation with a VBA project
  (Module1, the class CDeck, and three helper modules). pyOpenVBA cannot create
  a .ppt, so this one is copied from XLIDE's
  `tests/fixtures/binaries/PowerPointFixture.ppt` (xlide_vscode e56098b, MIT,
  same author).
