OutFile "StarkBankInstaller.exe"

  !define NAME "StarkBankExcel"
  !define VERSION "2.0.3"
  !define SLUG "${NAME} v${VERSION}"

  Name "${NAME}"
  RequestExecutionLevel admin

Section

    SetOutPath "$EXEDIR\StarkBankExcel"

    ExecShell "" "explorer" "$EXEDIR\StarkBankExcel"

    File /r "C:\StarkExcelInstaller\*.*"
    
SectionEnd

Section

    ExecWait 'cmd.exe /c "certutil -addstore "TrustedPublisher" $EXEDIR\StarkBankExcel\StarkBank.cer"'

SectionEnd