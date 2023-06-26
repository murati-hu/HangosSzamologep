

!define MUI_PRODUCT "Hangos sz�mol�g�p" ;Define your own software name here
!define MUI_VERSION "" ;Define your own software version here

!include "MUI.nsh"

;--------------------------------
;Configuration

  OutFile "hsz_setup.exe"

  ShowInstDetails show

  InstallDir "$PROGRAMFILES\Hang_Szamologep"
  
  InstallDirRegKey HKCU "Software\Hang_Szamologep" ""

;--------------------------------
;Modern UI Configuration

  ;!define USE_SUBSECTION
  ;!define MUI_LICENSEPAGE
  !define MUI_COMPONENTSPAGE
  !define MUI_COMPONENTSPAGE_SMALLDESC
  !define MUI_DIRECTORYPAGE
  
  !define MUI_ABORTWARNING
  
  !define MUI_UNINSTALLER
  !define MUI_UNCONFIRMPAGE

  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"
  
;--------------------------------
;Language Strings

  ;Description
	LangString DESC_hsz ${LANG_HUNGARIAN} "Hangos sz�mol�g�p telep�t�se az �n sz�m�t�g�p�re."
	LangString DESC_VB6 ${LANG_HUNGARIAN} "Fut�shoz sz�ks�ges Visual Basic 6.0 (SP5) Runtime f�jlok telep�t�se.(XP alatt nem sz�ks�ges)"
	LangString DESC_dx7 ${LANG_HUNGARIAN} "DirectX 7 telep�t�se a hangok kezel�s�hez."
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Elt�vol�t� alkalmaz�s telep�t�se. (Uninstall)"

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

Section "Hangos sz�mol�g�p" hsz
	SectionIn RO

	detailprint ">>> Hangos sz�mol�g�p telep�t�se..."
	CreateDirectory "$INSTDIR"
	CopyFiles "$EXEDIR\hang_szam.exe" "$INSTDIR\" 64
  	CreateDirectory "$SMPROGRAMS\Hangos sz�mol�g�p"
	CreateShortCut "$SMPROGRAMS\Hangos sz�mol�g�p\Hangos sz�mol�g�p.lnk" "$INSTDIR\hang_szam.exe"
	detailprint ""

	detailprint ">>> Hangok m�sol�sa..."
	CreateDirectory "$INSTDIR\hangok"
	CopyFiles "$EXEDIR\hangok\*.*" "$INSTDIR\hangok" 2460
	detailprint ""	
SectionEnd

section "Microsoft Visual Basic 6.0 Runtime (SP5)" VB6
	detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telep�t�se..."
	execwait "$EXEDIR\vbrun.exe /q"
	detailprint ""
sectionend

section "DirectX 7" dx7
	detailprint ">>> Microsoft DirectX 7.0 telep�t�se..."
	execwait "$EXEDIR\dx70eng.exe /q"
	detailprint ""
sectionend

Section "Elt�vol�t� alkalmaz�s" Eltavolit
	detailprint ">>> Elt�vo�t� alkalmaz�s telep�t�se..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\Hangos sz�mol�g�p\Elt�vol�t�s.lnk" "$INSTDIR\eltavolit.exe" 
Sectionend 


!insertmacro MUI_SECTIONS_FINISHHEADER


!insertmacro MUI_FUNCTIONS_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${hsz} $(DESC_hsz)
	!insertmacro MUI_DESCRIPTION_TEXT ${VB6} $(DESC_VB6)
	!insertmacro MUI_DESCRIPTION_TEXT ${dx7} $(DESC_dx7)
	!insertmacro MUI_DESCRIPTION_TEXT ${Eltavolit} $(DESC_Eltavolit)
!insertmacro MUI_FUNCTIONS_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"
	delete "$INSTDIR\hangok\*.*"
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\Hangos sz�mol�g�p\*.*"
	rmdir "$SMPROGRAMS\Hangos sz�mol�g�p"
	rmdir "$INSTDIR\hangok"
	rmdir "$INSTDIR"

  	!insertmacro MUI_UNFINISHHEADER
SectionEnd