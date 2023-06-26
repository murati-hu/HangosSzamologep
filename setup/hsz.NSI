!include "MUI.nsh"

;--------------------------------
;Configuration

  OutFile "hsz_setup.exe"

  ShowInstDetails show

  InstallDir "$PROGRAMFILES\Hangos_Szamologep"

  InstallDirRegKey HKCU "Software\Hangos_Szamologep" ""

;--------------------------------
;Modern UI Configuration
  ;!define MUI_LICENSEPAGE
  !define MUI_COMPONENTSPAGE
  !define MUI_COMPONENTSPAGE_SMALLDESC
  !define MUI_DIRECTORYPAGE

  !define MUI_ABORTWARNING

  !define MUI_UNINSTALLER
  !define MUI_UNCONFIRMPAGE

;Pages

  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES

  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------
;Languages

  !insertmacro MUI_LANGUAGE "Hungarian"

;--------------------------------
;Language Strings

  ;Description
	LangString DESC_hsz ${LANG_HUNGARIAN} "Hangos számológép telepítése az ön számítógépére."
	LangString DESC_VB6 ${LANG_HUNGARIAN} "Futáshoz szükséges Visual Basic 6.0 (SP5) Runtime fájlok telepítése.(XP alatt nem szükséges)"
	LangString DESC_dx7 ${LANG_HUNGARIAN} "DirectX 7 telepítése a hangok kezeléséhez."
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Eltávolító alkalmazás telepítése. (Uninstall)"

;--------------------------------
;Data

  ;LicenseData "license.rtf"

;--------------------------------
;Reserve Files

  ;Things that need to be extracted on first (keep these lines before any File command!)
  ;Only useful for BZIP2 compression

  ;ReserveFile "${NSISDIR}\Contrib\Icons\modern-header.bmp"

;--------------------------------
;Installer Sections

Section "Hangos számológép" hsz
	SectionIn RO

	detailprint ">>> Hangos számológép telepítése..."
	;CreateDirectory "$INSTDIR"
	SetOutPath "$INSTDIR"
	File "..\hang_szam.exe"
	;CopyFiles "$EXEDIR\hang_szam.exe" "$INSTDIR\" 64
  	CreateDirectory "$SMPROGRAMS\Hangos számológép"
	CreateShortCut "$SMPROGRAMS\Hangos számológép\Hangos számológép.lnk" "$INSTDIR\hang_szam.exe"
	detailprint ""

	detailprint ">>> Hangok másolása..."
	CreateDirectory "$INSTDIR\hangok"
	SetOutPath "$INSTDIR\hangok"
	file "..\hangok\*.*"
	detailprint ""
SectionEnd

section "Microsoft Visual Basic 6.0 Runtime (SP5)" VB6
	detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telepítése..."
	SetOutPath "$INSTDIR"
	file "vbrun.exe"
	execwait "$INSTDIR\vbrun.exe /q"
	detailprint ""
sectionend

section "DirectX 7" dx7
	detailprint ">>> Microsoft DirectX 7.0 telepítése..."
	SetOutPath "$INSTDIR"
	file "dx70eng.exe"
	execwait "$INSTDIR\dx70eng.exe /q"

	SetOutPath "$SYSDIR"
	file "dx7vb.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/dx7vb.dll"
	detailprint ""
sectionend

Section "Eltávolító alkalmazás" Eltavolit
	detailprint ">>> Eltávoító alkalmazás telepítése..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\Hangos számológép\Eltávolítás.lnk" "$INSTDIR\eltavolit.exe"
Sectionend


!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${hsz} $(DESC_hsz)
	!insertmacro MUI_DESCRIPTION_TEXT ${VB6} $(DESC_VB6)
	!insertmacro MUI_DESCRIPTION_TEXT ${dx7} $(DESC_dx7)
	!insertmacro MUI_DESCRIPTION_TEXT ${Eltavolit} $(DESC_Eltavolit)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

Section "Uninstall"
	delete "$INSTDIR\hangok\*.*"
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\Hangos számológép\*.*"
	rmdir "$SMPROGRAMS\Hangos számológép"
	rmdir "$INSTDIR\hangok"
	rmdir "$INSTDIR"
SectionEnd