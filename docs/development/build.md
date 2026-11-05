# Developer build

The development build pins .NET SDK 10.0.403, ClosedXML 0.105.0 and Inno Setup 6.4.3. Run `build/build.ps1` with `-Dotnet` and `-Iscc` pointing to the SDK and that compiler. A self-contained Windows x64 executable and installer are generated under `artifacts`, which is excluded from Git. End-user documentation and license notices accompany the executable.

The installer must include compiled runtime resources and no business source files. Inno Setup 6.4.3 explicitly permits commercial use in its [original license](https://github.com/jrsoftware/issrc/blob/is-6_4_3/license.txt). Its terms differ from newer compiler releases; do not silently upgrade the compiler. Record the actual compiler version with the release.

Only synthetic fixtures belong in this public repository. Real mappings, accounts, balances, transactions and recordings stay local. Application source, development tools, evidence and production data are separate assets.

Business acceptance runs through the installed graphical application. Internal tests assist development and do not replace installed GUI acceptance or actual Excel editing checks. Formula cells carry engine-calculated numeric values in the output package; Excel recalculates upon opening and editing.
