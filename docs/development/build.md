# Developer build

The development build pins .NET SDK 10.0.403 and ClosedXML 0.105.0. Run `build/build.ps1` with `-Dotnet` and `-Iscc` pointing to the SDK and Inno Setup compiler. A self-contained Windows x64 executable and installer are generated under `artifacts`, which is excluded from Git. Only README.md is copied as end-user documentation.

The installer must include compiled runtime resources and no business source files. Select an Inno Setup compiler licensed for the intended distribution. Record the actual compiler version with the release.

Only synthetic fixtures belong in this public repository. Real mappings, accounts, balances, transactions and recordings stay local. Application source, development tools, evidence and production data are separate assets.

Business acceptance runs through the installed graphical application. Internal tests assist development and do not replace installed GUI acceptance or actual Excel editing checks. Formula cells carry engine-calculated numeric values in the output package; Excel recalculates upon opening and editing.
