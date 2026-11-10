;=============================================================
;
; LSECTOOL Professional V6
;
; Canal Longitudinal Section Generator
;
; Version : 6.0
; Platform: AutoCAD 2007+
; Excel   : Microsoft Excel COM
;
;=============================================================

(vl-load-com)

;-------------------------------------------------------------
; Application Information
;-------------------------------------------------------------

(setq *LSEC-VERSION* "6.0")

(setq *LSEC-NAME*
      "LSECTOOL Professional V6")

;-------------------------------------------------------------
; Global Objects
;-------------------------------------------------------------

(setq *LSEC-EXCEL*    nil)
(setq *LSEC-WORKBOOK* nil)
(setq *LSEC-SHEET*    nil)
(setq *LSEC-DEBUG* nil)
(setq *LSEC-HEADERS* nil)
(setq *LSEC-LASTROW* 0)
;-------------------------------------------------------------
; Data Storage
;-------------------------------------------------------------

(setq *LSEC-DATA* nil)

(setq *LSEC-HEADERS* nil)

(setq *LSEC-CONFIG* nil)

(setq *LSEC-SEGMENTS* nil)

;-------------------------------------------------------------
; Current Drawing Information
;-------------------------------------------------------------

(setq *LSEC-DATUM* nil)

(setq *LSEC-MIN-ELEV* nil)

(setq *LSEC-MAX-ELEV* nil)

(setq *LSEC-MIN-CHAINAGE* nil)

(setq *LSEC-MAX-CHAINAGE* nil)

(princ
 (strcat
  "\n"
  *LSEC-NAME*
  " Loaded."
  " Version "
  *LSEC-VERSION*
 ))

(princ)

;=============================================================
;
; Part 1B - Professional Error Handler
;
;=============================================================

(setq *LSEC-OLDERROR* nil)

(defun LSEC:RestoreEnvironment ( )

    ;; Restore original error handler
    (if *LSEC-OLDERROR*
        (setq *error* *LSEC-OLDERROR*)
    )

    ;; Close any open undo group safely
    (if (= 8 (logand 8 (getvar "UNDOCTL")))
        (command "_.UNDO" "_END")
    )

    (princ)
)

(defun LSEC:Error (msg)

    ;; Ignore normal user cancel
    (if (not
            (member
                (strcase msg)
                '(
                    "FUNCTION CANCELLED"
                    "QUIT / EXIT ABORT"
                    "CONSOLE BREAK"
                 )
            )
        )

        (progn
            (prompt
                (strcat
                    "\n[LSECTOOL] ERROR: "
                    msg
                )
            )
        )
    )

    (LSEC:RestoreEnvironment)

    (princ)
)

(defun LSEC:StartCommand ( )

    ;; Save AutoCAD error handler
    (setq *LSEC-OLDERROR* *error*)

    ;; Install LSEC handler
    (setq *error* LSEC:Error)

    ;; Start UNDO group
    (command "_.UNDO" "_BEGIN")

    (princ)
)

(defun LSEC:EndCommand ( )

    ;; Restore environment
    (LSEC:RestoreEnvironment)

    (princ)
)

;=============================================================
;
; Part 1C - Utility Functions
;
;=============================================================

;-------------------------------------------------------------
; Safe String Conversion
;-------------------------------------------------------------
(defun LSEC:ToString (val)
  (cond
    ((null val) "")
    ((= (type val) 'STR) val)
    (T (vl-princ-to-string val))
  )
)

;-------------------------------------------------------------
; Safe Real Conversion
;-------------------------------------------------------------
(defun LSEC:ToReal (val / n)
  (cond
    ((numberp val) (float val))
    ((= (type val) 'STR)
      (setq n (distof val 2))
      (if n (float n) nil)
    )
    (T nil)
  )
)

;-------------------------------------------------------------
; Safe Integer Conversion
;-------------------------------------------------------------
(defun LSEC:ToInteger (val / n)
  (setq n (LSEC:ToReal val))
  (if n (fix n) nil)
)

;-------------------------------------------------------------
; Is Numeric?
;-------------------------------------------------------------
(defun LSEC:IsNumber (val)
  (if (numberp (LSEC:ToReal val))
    T
    nil
  )
)

;-------------------------------------------------------------
; Trim String
;-------------------------------------------------------------
(defun LSEC:Trim (str)
  (if (= (type str) 'STR)
    (vl-string-trim " \t\r\n" str)
    ""
  )
)

;-------------------------------------------------------------
; Convert to Uppercase
;-------------------------------------------------------------
(defun LSEC:Upper (str)
  (strcase (LSEC:Trim str))
)

;-------------------------------------------------------------
; Safe Association Lookup
;-------------------------------------------------------------
(defun LSEC:AssocValue (key alist)
  (cdr (assoc key alist))
)

;-------------------------------------------------------------
; Layer Exists?
;-------------------------------------------------------------
(defun LSEC:LayerExists (lay)
  (if (tblsearch "LAYER" lay)
    T
    nil
  )
)

;-------------------------------------------------------------
; Text Style Exists?
;-------------------------------------------------------------
(defun LSEC:StyleExists (sty)
  (if (tblsearch "STYLE" sty)
    T
    nil
  )
)

;-------------------------------------------------------------
; Block Exists?
;-------------------------------------------------------------
(defun LSEC:BlockExists (blk)
  (if (tblsearch "BLOCK" blk)
    T
    nil
  )
)

;-------------------------------------------------------------
; Print LSEC Message
;-------------------------------------------------------------
(defun LSEC:Print (msg)
  (prompt (strcat "\n[LSECTOOL] " (LSEC:ToString msg)))
)

;-------------------------------------------------------------
; Debug Print
;-------------------------------------------------------------
(defun LSEC:Debug (msg)
  (if *LSEC-DEBUG*
    (LSEC:Print msg)
  )
)

;-------------------------------------------------------------
; Yes/No Prompt
;-------------------------------------------------------------
(defun LSEC:GetYesNo (msg default / ans)
  (initget "Yes No")
  (setq ans
        (getkword
          (strcat "\n"
                  msg
                  " [Yes/No] <"
                  default
                  ">: ")))
  (if ans ans default)
)

;-------------------------------------------------------------
; Pause
;-------------------------------------------------------------
(defun LSEC:Pause ()
  (grread T)
)

(princ "\nPart 1C - Utility Functions Loaded.")
(princ)

;-------------------------------------------------------------
; Normalize Excel Header
;-------------------------------------------------------------
(defun LSEC:NormalizeHeader (txt)

  (if txt

    (progn

      ;; Replace Excel line breaks with spaces
      (setq txt (vl-string-subst " " "\n" txt))
      (setq txt (vl-string-subst " " "\r" txt))

      ;; Trim spaces
      (setq txt (vl-string-trim " " txt))

      ;; Convert to uppercase
      (strcase txt)

    )

    ""

  )

)


;=============================================================
;
; Part 1D - Layer Manager
;
;=============================================================


;-------------------------------------------------------------
; Create Layer
;-------------------------------------------------------------

(defun LSEC:CreateLayer (lname color)

  (if (not (LSEC:LayerExists lname))

    (progn

      (command "_.-LAYER"
               "_M"
               lname
               "_C"
               color
               ""
               "")

      (LSEC:Print
        (strcat "Layer created: " lname))

    )

  )

)


;-------------------------------------------------------------
; Set Current Layer
;-------------------------------------------------------------

(defun LSEC:SetLayer (lname)

  (if (LSEC:LayerExists lname)

    (setvar "CLAYER" lname)

    (LSEC:Print
      (strcat "Layer not found: " lname))

  )

)


;-------------------------------------------------------------
; Initialize All LSEC Layers
;-------------------------------------------------------------

(defun LSEC:InitLayers ()

  (LSEC:Print "Initializing Layers...")


  (LSEC:CreateLayer "LSEC-EGL" "3")

  (LSEC:CreateLayer "LSEC-BED" "5")

  (LSEC:CreateLayer "LSEC-FSL" "4")

  (LSEC:CreateLayer "LSEC-BANK" "2")

  (LSEC:CreateLayer "LSEC-GRID" "7")

  (LSEC:CreateLayer "LSEC-TEXT" "7")

  (LSEC:CreateLayer "LSEC-DATUM" "1")

  (LSEC:CreateLayer "LSEC-STRUCT" "6")


  (LSEC:Print "Layer initialization complete")

  T

)

;=============================================================
;
; Part 1E - Text Style Manager (V6 Revised)
;
;=============================================================


;-------------------------------------------------------------
; Check Text Style Exists
;-------------------------------------------------------------

(defun LSEC:TextStyleExists (sname)

  (tblsearch "STYLE" sname)

)


;-------------------------------------------------------------
; Create Text Style
;-------------------------------------------------------------

(defun LSEC:CreateTextStyle (sname font)

  (if (not (LSEC:TextStyleExists sname))

    (progn

      (entmake
        (list
          '(0 . "STYLE")
          '(100 . "AcDbSymbolTableRecord")
          '(100 . "AcDbTextStyleTableRecord")
          (cons 2 sname)
          (cons 3 font)
          (cons 70 0)
          (cons 40 0.0)
          (cons 41 1.0)
          (cons 50 0.0)
          (cons 71 0)
        )
      )


      (LSEC:Print
        (strcat "Text style created: " sname))

    )

    (LSEC:Print
      (strcat "Text style exists: " sname))

  )

)



;-------------------------------------------------------------
; Initialize Text Styles
;-------------------------------------------------------------

(defun LSEC:InitTextStyles ()

  (LSEC:Print "Initializing Text Styles...")


  (LSEC:CreateTextStyle
    "LSEC-TEXT"
    "Tahoma")


  (LSEC:CreateTextStyle
    "LSEC-TITLE"
    "Tahoma")


  (LSEC:CreateTextStyle
    "LSEC-ANNOT"
    "Tahoma")


  ; Set current style

  (if (tblsearch "STYLE" "LSEC-TEXT")

    (setvar "TEXTSTYLE" "LSEC-TEXT")

  )


  (LSEC:Print
    "Text styles initialized")


  T

)

;=============================================================
;
; Part 2A - Configuration Initialization
;
;=============================================================


;-------------------------------------------------------------
; Global configuration storage
;
; IMPORTANT:
; Do not overwrite an existing configuration every time
; InitConfig is called. DrawCompleteSection calls InitConfig
; during every split-section pass.
;
;-------------------------------------------------------------

(if (not (boundp '*LSEC-CONFIG*))
  (setq *LSEC-CONFIG* nil)
)


;-------------------------------------------------------------
; Initialize Configuration
;
; Adds default values only when they do not already exist.
; Existing user settings are preserved.
;
;-------------------------------------------------------------

(defun LSEC:InitConfig (/)

  ;-----------------------------------------------------------
  ; Create configuration storage if necessary
  ;-----------------------------------------------------------

  (if (not *LSEC-CONFIG*)
    (setq *LSEC-CONFIG* nil)
  )


  ;-----------------------------------------------------------
  ; Add default values only if missing
  ;-----------------------------------------------------------

  (if (not (assoc "HSCALE" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "HSCALE" 1000.0))
      )
    )
  )

  (if (not (assoc "VEXAG" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "VEXAG" 10.0))
      )
    )
  )

  (if (not (assoc "DATUM" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "DATUM" nil))
      )
    )
  )

  (if (not (assoc "TEXTHEIGHT" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "TEXTHEIGHT" 3.5))
      )
    )
  )

  (if (not (assoc "TITLEHEIGHT" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "TITLEHEIGHT" 4.0))
      )
    )
  )

  (if (not (assoc "SPLITMODE" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "SPLITMODE" 0))
      )
    )
  )

  (if (not (assoc "SPLITLENGTH" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "SPLITLENGTH" 500.0))
      )
    )
  )

  (if (not (assoc "SECTIONGAP" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "SECTIONGAP" 50.0))
      )
    )
  )

  (if (not (assoc "DRAWGRID" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "DRAWGRID" T))
      )
    )
  )

  (if (not (assoc "DRAWTITLE" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "DRAWTITLE" T))
      )
    )
  )

  (if (not (assoc "GRIDX" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "GRIDX" 20.0))
      )
    )
  )

  (if (not (assoc "GRIDY" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (append *LSEC-CONFIG*
        (list (cons "GRIDY" 1.0))
      )
    )
  )


  (LSEC:Print
    "Configuration initialized/preserved"
  )

  *LSEC-CONFIG*

)

;=============================================================
;
; Part 2B - Configuration Getter
;
;=============================================================


;-------------------------------------------------------------
; Get Configuration Value
;-------------------------------------------------------------

(defun LSEC:GetConfig (key)

  (if *LSEC-CONFIG*

    (progn

      (setq item
        (assoc key *LSEC-CONFIG*)
      )


      (if item

        (cdr item)

        nil

      )

    )

    nil

  )

)

;=============================================================
;
; Part 2C - Configuration Setter
;
;=============================================================


;-------------------------------------------------------------
; Set Configuration Value
;-------------------------------------------------------------

(defun LSEC:SetConfig (key value)

  ; If configuration does not exist, initialize it

  (if (not *LSEC-CONFIG*)

    (LSEC:InitConfig)

  )


  ; Check if key already exists

  (setq item
    (assoc key *LSEC-CONFIG*)
  )


  (if item

    ; Replace existing value

    (setq *LSEC-CONFIG*

      (subst
        (cons key value)
        item
        *LSEC-CONFIG*
      )

    )


    ; Add new configuration item

    (setq *LSEC-CONFIG*

      (append
        *LSEC-CONFIG*
        (list (cons key value))
      )

    )

  )


  (LSEC:Print
    (strcat
      "Config updated: "
      key
    )
  )


  value

)

;=============================================================
;
; Part 2D - Drawing Configuration Input
;
;=============================================================


;-------------------------------------------------------------
; Get Drawing Configuration
;-------------------------------------------------------------

(defun LSEC:GetDrawingConfig ()

  (LSEC:Print
    "Getting drawing configuration...")


  ; Ensure configuration exists

  (if (not *LSEC-CONFIG*)

    (LSEC:InitConfig)

  )


  ;-----------------------------------------------------------
  ; Horizontal Scale
  ;-----------------------------------------------------------

  (setq hscale

    (getreal
      (strcat
        "\nEnter horizontal scale denominator <"
        (rtos (LSEC:GetConfig "HSCALE") 2 0)
        ">: "
      )
    )

  )


  (if hscale

    (LSEC:SetConfig "HSCALE" hscale)

  )



  ;-----------------------------------------------------------
  ; Vertical Exaggeration
  ;-----------------------------------------------------------

  (setq vexag

    (getreal
      (strcat
        "\nEnter vertical exaggeration <"
        (rtos (LSEC:GetConfig "VEXAG") 2 1)
        ">: "
      )
    )

  )


  (if vexag

    (LSEC:SetConfig "VEXAG" vexag)

  )


  ;-----------------------------------------------------------
  ; Datum Level
  ;-----------------------------------------------------------

  (setq datum

    (getreal
      "\nEnter datum level <Automatic>: "
    )

  )


  (if datum

    (LSEC:SetConfig "DATUM" datum)

  )



  ;-----------------------------------------------------------
  ; Split Mode
  ;-----------------------------------------------------------

  (setq splitmode

    (getint
      (strcat
        "\nSplit drawing? 0=No, 1=Yes <"
        (itoa (LSEC:GetConfig "SPLITMODE"))
        ">: "
      )
    )

  )


  (if splitmode

    (LSEC:SetConfig "SPLITMODE" splitmode)

  )



  ;-----------------------------------------------------------
  ; Split Length
  ;-----------------------------------------------------------

  (if (= (LSEC:GetConfig "SPLITMODE") 1)

    (progn

      (setq splitlength

        (getreal
          (strcat
            "\nEnter split length <"
            (rtos
              (LSEC:GetConfig "SPLITLENGTH")
              2
              0
            )
            ">: "
          )
        )

      )


      (if splitlength

        (LSEC:SetConfig
          "SPLITLENGTH"
          splitlength)

      )

    )

  )


  (LSEC:Print
    "Drawing configuration completed")


  *LSEC-CONFIG*

)

;=============================================================
;
; Part 3A - Excel COM Manager
;
;=============================================================


;-------------------------------------------------------------
; Start Excel Application
;-------------------------------------------------------------

(defun LSEC:ExcelStart ()

  (if *LSEC-EXCEL*

    (progn

      (LSEC:Print
        "Excel already running")

      *LSEC-EXCEL*

    )


    (progn

      (setq *LSEC-EXCEL*

        (vlax-create-object "Excel.Application")

      )


      (if *LSEC-EXCEL*

        (progn

          ; Keep Excel hidden

          (vlax-put-property
            *LSEC-EXCEL*
            'Visible
            :vlax-false
          )


          (LSEC:Print
            "Excel started")

        )


        (LSEC:Print
          "Excel start failed")

      )


      *LSEC-EXCEL*

    )

  )

)



;-------------------------------------------------------------
; Release COM Object
;-------------------------------------------------------------

(defun LSEC:ReleaseObject (obj)

  (if obj

    (progn

      (vl-catch-all-apply
        'vlax-release-object
        (list obj)
      )

    )

  )

)



;-------------------------------------------------------------
; Close Excel Safely
;-------------------------------------------------------------

(defun LSEC:ExcelClose ()

  (if *LSEC-WORKBOOK*

    (progn

      (vl-catch-all-apply
        'vlax-invoke-method
        (list
          *LSEC-WORKBOOK*
          'Close
          :vlax-false
        )
      )

      (LSEC:ReleaseObject
        *LSEC-WORKBOOK*)

      (setq *LSEC-WORKBOOK* nil)

    )

  )



  (if *LSEC-EXCEL*

    (progn

      (vl-catch-all-apply
        'vlax-invoke-method
        (list
          *LSEC-EXCEL*
          'Quit
        )
      )


      (LSEC:ReleaseObject
        *LSEC-EXCEL*)


      (setq *LSEC-EXCEL* nil)

    )

  )


  (LSEC:Print
    "Excel closed")


  T

)

;=============================================================
;
; Part 3B - Workbook Manager
;
;=============================================================



;-------------------------------------------------------------
; Open Workbook
;-------------------------------------------------------------

(defun LSEC:WorkbookOpen (filepath)

  (LSEC:Print
    "Opening workbook...")


  ; Start Excel if not running

  (if (not *LSEC-EXCEL*)

    (LSEC:ExcelStart)

  )


  ; Check file exists

  (if (not (findfile filepath))

    (progn

      (LSEC:Print
        "Workbook file not found")

      nil

    )


    (progn


      ; Open workbook

      (setq *LSEC-WORKBOOK*

        (vlax-invoke-method

          (vlax-get-property
            *LSEC-EXCEL*
            'Workbooks)

          'Open

          filepath

        )

      )


      (if *LSEC-WORKBOOK*

        (LSEC:Print
          "Workbook opened")

        (LSEC:Print
          "Workbook open failed")

      )


      *LSEC-WORKBOOK*

    )

  )

)



;-------------------------------------------------------------
; Get Active Workbook
;-------------------------------------------------------------

(defun LSEC:GetWorkbook ()

  *LSEC-WORKBOOK*

)



;-------------------------------------------------------------
; Close Workbook Only
;-------------------------------------------------------------

(defun LSEC:WorkbookClose ()

  (if *LSEC-WORKBOOK*

    (progn


      (vl-catch-all-apply

        'vlax-invoke-method

        (list

          *LSEC-WORKBOOK*

          'Close

          :vlax-false

        )

      )


      (vlax-release-object
        *LSEC-WORKBOOK*
      )


      (setq *LSEC-WORKBOOK* nil)


      (LSEC:Print
        "Workbook closed")


    )

  )


  T

)

;=============================================================
;
; Part 3C - Worksheet Manager (V6 Fixed)
;
;=============================================================


;-------------------------------------------------------------
; Get Worksheet By Index or Name
;-------------------------------------------------------------

(defun LSEC:GetWorksheet (sheet)

  (LSEC:Print
    "Getting worksheet...")


  (if (not *LSEC-WORKBOOK*)

    (progn

      (LSEC:Print
        "No workbook available")

      nil

    )


    (progn


      (setq sheets

        (vlax-get-property
          *LSEC-WORKBOOK*
          'Worksheets
        )

      )


      (setq result

        (vl-catch-all-apply

          'vlax-get-property

          (list
            sheets
            'Item
            sheet
          )

        )

      )


      (if
        (vl-catch-all-error-p result)

        (progn

          (LSEC:Print
            "Worksheet selection failed")

          (setq *LSEC-SHEET* nil)

        )


        (progn

          (setq *LSEC-SHEET* result)

          (LSEC:Print
            "Worksheet selected")

        )

      )


      (vlax-release-object sheets)


      *LSEC-SHEET*

    )

  )

)




;-------------------------------------------------------------
; Prompt user to select a worksheet by number
; Returns the selected worksheet object, or nil on cancel.
; Compatible with older AutoCAD/Excel COM interfaces.
;-------------------------------------------------------------

(defun LSEC:PromptWorksheetSelection
       (/ sheets countResult sheetResult sheetName
          index choice selected)

  (if (not *LSEC-WORKBOOK*)
    (progn
      (LSEC:Print "No Excel workbook is available.")
      nil
    )
    (progn
      (setq sheets
        (vl-catch-all-apply
          'vlax-get-property
          (list *LSEC-WORKBOOK* 'Worksheets)
        )
      )

      (if (vl-catch-all-error-p sheets)
        (progn
          (LSEC:Print "Could not retrieve the workbook worksheets.")
          nil
        )
        (progn
          (setq countResult
            (vl-catch-all-apply
              'vlax-get-property
              (list sheets 'Count)
            )
          )

          (if
            (or
              (vl-catch-all-error-p countResult)
              (not (numberp countResult))
              (< countResult 1)
            )
            (progn
              (LSEC:Print "No worksheets were found in this workbook.")
              (vl-catch-all-apply 'vlax-release-object (list sheets))
              nil
            )
            (progn
              (princ "\n")
              (LSEC:Print "Available Excel worksheets:")
              (setq index 1)
              (while (<= index countResult)
                (setq sheetResult
                  (vl-catch-all-apply
                    'vlax-get-property
                    (list sheets 'Item index)
                  )
                )
                (if (not (vl-catch-all-error-p sheetResult))
                  (progn
                    (setq sheetName
                      (vl-catch-all-apply
                        'vlax-get-property
                        (list sheetResult 'Name)
                      )
                    )
                    (if (vl-catch-all-error-p sheetName)
                      (setq sheetName "<Name unavailable>")
                    )
                    (LSEC:Print
                      (strcat
                        (itoa index)
                        ". "
                        sheetName
                      )
                    )
                    (vl-catch-all-apply
                      'vlax-release-object
                      (list sheetResult)
                    )
                  )
                  (LSEC:Print
                    (strcat (itoa index) ". <Could not read worksheet name>")
                  )
                )
                (setq index (1+ index))
              )

              (setq choice nil)
              (while
                (and
                  (not selected)
                  (not choice)
                )
                (setq choice
                  (getint
                    (strcat
                      "\nEnter worksheet number (1-"
                      (itoa countResult)
                      ") or press Enter to cancel: "
                    )
                  )
                )
                (cond
                  ((null choice)
                    (LSEC:Print "Worksheet selection cancelled.")
                  )
                  ((and (>= choice 1) (<= choice countResult))
                    (setq selected T)
                  )
                  (T
                    (princ
                      (strcat
                        "\nInvalid selection. Enter a number from 1 to "
                        (itoa countResult)
                        "."
                      )
                    )
                    (setq choice nil)
                  )
                )
              )

              (vl-catch-all-apply 'vlax-release-object (list sheets))

              (if selected
                (LSEC:GetWorksheet choice)
                nil
              )
            )
          )
        )
      )
    )
  )
)


;-------------------------------------------------------------
; Get Active Worksheet
;-------------------------------------------------------------

(defun LSEC:GetActiveSheet ()

  *LSEC-SHEET*

)



;-------------------------------------------------------------
; Release Worksheet
;-------------------------------------------------------------

(defun LSEC:ReleaseWorksheet ()

  (if *LSEC-SHEET*

    (progn

      (vlax-release-object
        *LSEC-SHEET*
      )

      (setq *LSEC-SHEET* nil)

      (LSEC:Print
        "Worksheet released")

    )

  )

  T

)

;=============================================================
;
; Part 3D - Cell Reader
;
;=============================================================


;-------------------------------------------------------------
; Convert row/column to Excel address
;-------------------------------------------------------------

(defun LSEC:ExcelAddress (row col / letters n)

  (setq letters "")

  (while (> col 0)

    (setq n
      (rem (- col 1) 26)
    )

    (setq letters
      (strcat
        (chr (+ 65 n))
        letters
      )
    )

    (setq col
      (fix (/ (- col 1) 26))
    )

  )

  (strcat
    letters
    (itoa row)
  )

)



;-------------------------------------------------------------
; Read Excel Cell
;-------------------------------------------------------------

(defun LSEC:ReadCell (row col / address range value)

  (if (not *LSEC-SHEET*)

    nil

    (progn

      (setq address
        (LSEC:ExcelAddress row col)
      )


      (setq range

        (vlax-get-property
          *LSEC-SHEET*
          'Range
          address
        )

      )


      (setq value

        (vlax-get-property
          range
          'Value2
        )

      )


      (if (= (type value) 'VARIANT)

        (setq value
          (vlax-variant-value value)
        )

      )


      (vlax-release-object range)


      value

    )

  )

)



(defun LSEC:ReadText (row col)

  (LSEC:ToString
    (LSEC:ReadCell row col)
  )

)



(defun LSEC:ReadReal (row col)

  (LSEC:ToReal
    (LSEC:ReadCell row col)
  )

)

;-------------------------------------------------------------
; Get Header Column
;-------------------------------------------------------------

(defun LSEC:GetHeaderColumn
       (header / canonical item)

  (setq canonical
        (LSEC:CanonicalHeader header))

  (setq item
        (assoc canonical *LSEC-HEADERS*))

  (if item
    (cdr item)
    nil
  )

)

;-------------------------------------------------------------
; Read Excel Headers
;-------------------------------------------------------------
(defun LSEC:FindHeaders (/ col txt done)

  (setq *LSEC-HEADERS* nil)

  (setq col 1)
  (setq done nil)

  (while (not done)

    (setq txt
  (LSEC:NormalizeHeader
    (LSEC:ReadText 1 col)
  )
)

    (if (= txt "")

      (setq done T)

      (progn

        (setq *LSEC-HEADERS*
          (append
            *LSEC-HEADERS*
            (list
              (cons
                (LSEC:CanonicalHeader txt)
                col
              )
            )
          )
        )

        (setq col (1+ col))

      )

    )

  )

 (if *LSEC-DEBUG*
  (LSEC:Print
    (strcat
      "Headers found: "
      (itoa (length *LSEC-HEADERS*))
    )
  )
)

) ; <-- ADD THIS LINE. Closes LSEC:FindHeaders

;-------------------------------------------------------------
; Print Header Map
;-------------------------------------------------------------

(defun LSEC:PrintHeaders (/)

  (foreach item *LSEC-HEADERS*

    (princ
      (strcat
        "\n"
        (car item)
        " -> "
        (itoa (cdr item))
      )
    )

  )

  (princ)

)

;-------------------------------------------------------------
; Header Aliases
;-------------------------------------------------------------

(setq *LSEC-HEADER-ALIASES*

 '(
    ("CH"               . "CH")
    ("CHAINAGE"         . "CH")
    ("CHAINAGE (M)"     . "CH")

    ("O.G.L"            . "O.G.L")
    ("OGL"              . "O.G.L")
    ("GROUND LEVEL"     . "O.G.L")

    ("C.B.L"            . "C.B.L")
    ("CBL"              . "C.B.L")
    ("BED LEVEL"        . "C.B.L")

    ("F.S.L"            . "F.S.L")
    ("FSL"              . "F.S.L")

    ("T.B.L"            . "T.B.L")
    ("TBL"              . "T.B.L")

    ("CUT"              . "CUT")

    ("FILL"             . "FILL")

    ("STRUCTURE NAME"   . "STRUCTURE NAME")
    ("STRUCTURE"        . "STRUCTURE NAME")
 )

)

;-------------------------------------------------------------
; Normalize Header Alias
;-------------------------------------------------------------

(defun LSEC:CanonicalHeader (header / item)

  (setq header
        (LSEC:NormalizeHeader header))

  (setq item
        (assoc header *LSEC-HEADER-ALIASES*))

  (if item

      (cdr item)

      header

  )

)

;-------------------------------------------------------------
; Header Exists
;-------------------------------------------------------------

(defun LSEC:HeaderExists (header / found)

  (setq found nil)

  (foreach item *LSEC-HEADERS*

    (if (= (car item)
           (LSEC:CanonicalHeader header))

      (setq found T)

    )

  )

  found

)

;-------------------------------------------------------------
; Validate Required Headers
;-------------------------------------------------------------

(defun LSEC:ValidateHeaders (/ required ok)

  (setq required

   '(
      "CH"
      "O.G.L"
      "C.B.L"
      "F.S.L"
      "T.B.L"
      "CUT"
      "FILL"
      "STRUCTURE NAME"
    )
  )

  (setq ok T)

  (foreach hdr required

    (if (not (LSEC:HeaderExists hdr))

      (progn

        (LSEC:Print

          (strcat
            "Missing required header: "
            hdr
          )
        )

        (setq ok nil)

      )

    )

  )

  (if ok

      (LSEC:Print
        "Header validation successful."
      )

  )

  ok

)

;-------------------------------------------------------------
; Find Last Data Row
;-------------------------------------------------------------

(defun LSEC:FindLastRow (/ col row value)

  ;; Default

  (setq *LSEC-LASTROW* 0)

  ;; CH column

  (setq col

        (LSEC:GetHeaderColumn "CH")
  )

  (if (null col)

    (progn

      (LSEC:Print
        "CH column not found."
      )

      nil

    )

    (progn

      ;; Data starts after header

      (setq row 2)

  (while
  (LSEC:IsNumber
    (LSEC:ReadCell row col)
  )

  (setq *LSEC-LASTROW* row)
  (setq row (1+ row))
)

(LSEC:Print

  (strcat

    "Last data row: "

    (itoa *LSEC-LASTROW*)

  )

)

*LSEC-LASTROW*

)

)

)

;-------------------------------------------------------------
; Build Record
;-------------------------------------------------------------

(defun LSEC:BuildRecord
       (ch egl bed fsl bank cut fill struct)

  (list
    (cons "CHAINAGE" ch)
    (cons "EGL"       egl)
    (cons "BED"       bed)
    (cons "FSL"       fsl)
    (cons "BANK"      bank)
    (cons "CUT"       cut)
    (cons "FILL"      fill)
    (cons "STRUCT"    struct)
  )

)

;-------------------------------------------------------------
; Read One Excel Record
;-------------------------------------------------------------

(defun LSEC:ReadRecord (row / rec)

  (setq rec

    (LSEC:BuildRecord

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "CH"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "O.G.L"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "C.B.L"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "F.S.L"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "T.B.L"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "CUT"))

      (LSEC:ReadReal row (LSEC:GetHeaderColumn "FILL"))

      (LSEC:ReadText row (LSEC:GetHeaderColumn "STRUCTURE NAME"))

    )

  )

  rec

)

(LSEC:Print "Before Part 3E.4")

;-------------------------------------------------------------
; Print Record
;-------------------------------------------------------------

(defun LSEC:PrintRecord (rec)

  (foreach item rec

    (princ

      (strcat

        "\n"

        (car item)

        " = "

        (LSEC:ToString (cdr item))

      )

    )

  )

  (princ)

)

(LSEC:Print "After Part 3E.4")

(LSEC:Print "END OF FILE LOADED")

;=============================================================
;
; Part 3E.5 - Build LSEC Data
;
;=============================================================


;-------------------------------------------------------------
; Load All Excel Data
;-------------------------------------------------------------

(defun LSEC:LoadExcelData (/ row rec)

  (LSEC:Print "Loading Excel data...")

  (LSEC:FindHeaders)

  ; Clear previous data
  (setq *LSEC-DATA* nil)

  (if (<= *LSEC-LASTROW* 1)
    (LSEC:FindLastRow)
  )


  ; Check last row

  (if (<= *LSEC-LASTROW* 1)

    (progn

      (LSEC:Print
        "No Excel data available."
      )

      nil

    )


    (progn


      ; Start after header

      (setq row 2)


      (while (<= row *LSEC-LASTROW*)


        (setq rec

          (LSEC:ReadRecord row)

        )


        (if rec

          (setq *LSEC-DATA*

            (append
              *LSEC-DATA*
              (list rec)
            )

          )

        )


        (setq row (1+ row))


      )


      (LSEC:Print

        (strcat

          "Rows loaded: "

          (itoa (length *LSEC-DATA*))

        )

      )


      *LSEC-DATA*

    )

  )

)





;-------------------------------------------------------------
; Get Number of Loaded Records
;-------------------------------------------------------------

(defun LSEC:GetDataCount ()

  (length *LSEC-DATA*)

)



;-------------------------------------------------------------
; Get Record by Index
;-------------------------------------------------------------

(defun LSEC:GetDataRecord (index)

  (nth index *LSEC-DATA*)

)


;=============================================================
;
; Part 4A - Datum Calculation
;
;=============================================================


;-------------------------------------------------------------
; Find Minimum Profile Elevation
;-------------------------------------------------------------

(LSEC:Print "Before FindMinElevation")

(defun LSEC:FindMinElevation (/ rec minElev elev)


  (setq minElev nil)


  (foreach rec *LSEC-DATA*

    ; Check EGL

    (setq elev
      (cdr (assoc "EGL" rec))
    )

    (if elev

      (if (or (null minElev)
              (< elev minElev)
          )

        (setq minElev elev)

      )

    )


    ; Check BED

    (setq elev
      (cdr (assoc "BED" rec))
    )

    (if elev

      (if (or (null minElev)
              (< elev minElev)
          )

        (setq minElev elev)

      )

    )


    ; Check FSL

    (setq elev
      (cdr (assoc "FSL" rec))
    )

    (if elev

      (if (or (null minElev)
              (< elev minElev)
          )

        (setq minElev elev)

      )

    )


  )


  (setq *LSEC-MIN-ELEV* minElev)


  (LSEC:Print

    (strcat

      "Minimum elevation: "

      (rtos minElev 2 3)

    )

  )


  minElev

)

;-------------------------------------------------------------
; Calculate Drawing Datum
;-------------------------------------------------------------

(defun LSEC:CalculateDatum (/ minElev offset)

  (setq minElev
        (LSEC:FindMinElevation)
  )


  (if minElev

    (progn

      ; Clearance below lowest point

      (setq offset 2.0)


      (setq *LSEC-DATUM*
        (- minElev offset)
      )


      (LSEC:SetConfig
        "DATUM"
        *LSEC-DATUM*
      )


      (LSEC:Print

        (strcat

          "Datum calculated: "

          (rtos *LSEC-DATUM* 2 3)

        )

      )


      *LSEC-DATUM*

    )


    nil

  )

)

(LSEC:Print "After CalculateDatum")


;-------------------------------------------------------------
; Find Maximum Profile Elevation
;-------------------------------------------------------------

(defun LSEC:FindMaxElevation (/ rec maxElev elev)


  (setq maxElev nil)


  (foreach rec *LSEC-DATA*


    ; Check EGL

    (setq elev
      (cdr (assoc "EGL" rec))
    )

    (if elev

      (if (or (null maxElev)
              (> elev maxElev)
          )

        (setq maxElev elev)

      )

    )


    ; Check BED

    (setq elev
      (cdr (assoc "BED" rec))
    )

    (if elev

      (if (or (null maxElev)
              (> elev maxElev)
          )

        (setq maxElev elev)

      )

    )


    ; Check FSL

    (setq elev
      (cdr (assoc "FSL" rec))
    )

    (if elev

      (if (or (null maxElev)
              (> elev maxElev)
          )

        (setq maxElev elev)

      )

    )


    ; Check BANK

    (setq elev
      (cdr (assoc "BANK" rec))
    )

    (if elev

      (if (or (null maxElev)
              (> elev maxElev)
          )

        (setq maxElev elev)

      )

    )


  )


  (setq *LSEC-MAX-ELEV* maxElev)

(if maxElev

  (LSEC:Print
    (strcat
      "Maximum elevation: "
      (rtos maxElev 2 3)
    )
  )

  (LSEC:Print
    "Maximum elevation: no data found (*LSEC-DATA* is empty or has no elevation values)."
  )

)

maxElev

)

;-------------------------------------------------------------
; Calculate Elevation Range
;-------------------------------------------------------------

(defun LSEC:CalculateElevationRange
       (/ minElev maxElev range)


  ; Get minimum elevation

  (setq minElev
    (if *LSEC-MIN-ELEV*

      *LSEC-MIN-ELEV*

      (LSEC:FindMinElevation)

    )
  )


  ; Get maximum elevation

  (setq maxElev
    (if *LSEC-MAX-ELEV*

      *LSEC-MAX-ELEV*

      (LSEC:FindMaxElevation)

    )
  )


  ; Check values

  (if (and minElev maxElev)

    (progn


      (setq range
        (- maxElev minElev)
      )


      (setq *LSEC-ELEV-RANGE*
        range
      )


      (LSEC:Print

        (strcat

          "Elevation range: "

          (rtos range 2 3)

        )

      )


      range

    )


    nil

  )

)

;-------------------------------------------------------------
; Part 4B.1 - Initialize Horizontal Scale
;-------------------------------------------------------------

(defun LSEC:GetHorizontalScale (/ scale)

  (setq scale
    (LSEC:GetConfig "HSCALE")
  )


  (if scale

    (progn

      (setq *LSEC-HSCALE*
        scale
      )


      (LSEC:Print

        (strcat

          "Horizontal scale: 1:"

          (rtos scale 2 0)

        )

      )


      scale

    )

    nil

  )

)

;-------------------------------------------------------------
; Part 4B.1.1 - Convert Chainage to Drawing X
; Uses horizontal drawing scale.
;-------------------------------------------------------------

(defun LSEC:CHtoX (chainage)

  (if chainage
    chainage
    nil
  )

)

;-------------------------------------------------------------
; Part 4B.2 - Get Vertical Scale
;-------------------------------------------------------------

(defun LSEC:GetVerticalScale
       (/ vexag)

  (setq vexag
        (LSEC:GetConfig "VEXAG")
  )

  (if (and vexag (numberp vexag))

    (progn
      (if *LSEC-DEBUG*
        (LSEC:Print
          (strcat
            "Vertical scale factor: "
            (rtos vexag 2 6)
          )
        )
      )

      vexag
    )

    10.0

  )

)

;-------------------------------------------------------------
; Part 4B.2.1 - Convert Elevation to Drawing Y
;-------------------------------------------------------------

(defun LSEC:ELtoY (elevation / datum vscale y)


  ; Get datum

  (setq datum

    (if *LSEC-DATUM*

      *LSEC-DATUM*

      (LSEC:GetConfig "DATUM")

    )

  )


  ; Get vertical scale

  (setq vscale

    (if *LSEC-VSCALE*

      *LSEC-VSCALE*

      (LSEC:GetVerticalScale)

    )

  )


  (if (and elevation datum vscale)

    (progn

      (setq y

        (*

          (- elevation datum)

          vscale

        )

      )


      y

    )


    nil

  )

)

;-------------------------------------------------------------
; Part 4C.1 - Build Drawing Point
;-------------------------------------------------------------

(defun LSEC:BuildPoint (chainage elevation / x y)

  (if (and (numberp chainage)
           (numberp elevation))

    (progn

      ;-------------------------------------------------------
      ; X coordinate
      ;
      ; Use ActiveCHtoX so split sections are positioned at
      ; their own frame X location. For normal unsplit drawing,
      ; ActiveCHtoX behaves like the original CHtoX.
      ;-------------------------------------------------------

      (setq x
        (LSEC:ActiveCHtoX chainage)
      )

      ;-------------------------------------------------------
      ; Y coordinate
      ;-------------------------------------------------------

      (setq y
        (LSEC:ELtoY elevation)
      )

      (if (and (numberp x)
               (numberp y))

        (list x y 0.0)

        nil
      )
    )

    nil
  )
)

;-------------------------------------------------------------
; Part 4C.2 - Get EGL Point
;-------------------------------------------------------------

(defun LSEC:GetEGLPoint (rec / ch egl)

  (setq ch
        (cdr (assoc "CHAINAGE" rec)))

  (setq egl
        (cdr (assoc "EGL" rec)))

  (if (and (numberp ch)
           (numberp egl))

      (LSEC:BuildPoint ch egl)

      nil

  )

)

;-------------------------------------------------------------
; Part 4C.3 - Get BED Point
;-------------------------------------------------------------

(defun LSEC:GetBEDPoint (rec / ch bed)

  (setq ch
        (cdr (assoc "CHAINAGE" rec)))

  (setq bed
        (cdr (assoc "BED" rec)))

  (if (and (numberp ch)
           (numberp bed))

      (LSEC:BuildPoint ch bed)

      nil

  )

)

;-------------------------------------------------------------
; Part 4C.4 - Get FSL Point
;-------------------------------------------------------------

(defun LSEC:GetFSLPoint (rec / ch fsl)

  (setq ch
        (cdr (assoc "CHAINAGE" rec)))

  (setq fsl
        (cdr (assoc "FSL" rec)))

  (if (and (numberp ch)
           (numberp fsl))

      (LSEC:BuildPoint ch fsl)

      nil

  )

)

;-------------------------------------------------------------
; Part 4C.5 - Get BANK Point
;-------------------------------------------------------------

(defun LSEC:GetBANKPoint (rec / ch bank)

  (setq ch
        (cdr (assoc "CHAINAGE" rec)))

  (setq bank
        (cdr (assoc "BANK" rec)))

  (if (and (numberp ch)
           (numberp bank))

      (LSEC:BuildPoint ch bank)

      nil

  )

)

;-------------------------------------------------------------
; Part 4C.6.1 - Build Profile Point List
;-------------------------------------------------------------

(defun LSEC:BuildProfilePoints (PointFunction / pts pt)

  (setq pts nil)

  (foreach rec *LSEC-DATA*

    (setq pt
          (apply PointFunction (list rec)))

    (if pt
      (setq pts
            (append pts (list pt)))
    )

  )

  pts

)

;-------------------------------------------------------------
; Part 4C.6.2 - EGL Point List
;-------------------------------------------------------------

(defun LSEC:GetEGLPoints ()

  (LSEC:BuildProfilePoints
    'LSEC:GetEGLPoint)

)

;-------------------------------------------------------------
; Part 4C.6.3 - BED Point List
;-------------------------------------------------------------

(defun LSEC:GetBEDPoints ()

  (LSEC:BuildProfilePoints
    'LSEC:GetBEDPoint)

)

;-------------------------------------------------------------
; Part 4C.6.4 - FSL Point List
;-------------------------------------------------------------

(defun LSEC:GetFSLPoints ()

  (LSEC:BuildProfilePoints
    'LSEC:GetFSLPoint)

)

;-------------------------------------------------------------
; Part 4C.6.5 - BANK Point List
;-------------------------------------------------------------

(defun LSEC:GetBANKPoints ()

  (LSEC:BuildProfilePoints
    'LSEC:GetBANKPoint)

)

;=============================================================
;
; Part 5A.1 - LWPolyline Creator (Professional)
;
;=============================================================

(defun LSEC:DrawPolyline (pts layer / data ent)

  (if (and (listp pts)
           (> (length pts) 1)
           (LSEC:LayerExists layer))

    (progn

      (setq data

        (list

          '(0 . "LWPOLYLINE")

          '(100 . "AcDbEntity")

          (cons 8 layer)

          '(100 . "AcDbPolyline")

          (cons 90 (length pts))

          '(70 . 0)

        )

      )

      (foreach p pts

        (setq data

          (append

            data

            (list

              (cons 10

                (list

                  (car p)

                  (cadr p)

                )

              )

            )

          )

        )

      )

      (setq ent
      (entmakex data))

(if ent
    ent
    (progn
      (LSEC:Print "Polyline creation failed.")
      nil
    )
  )

)

)

)

;=============================================================
;
; Part 5A.2 - Text Creator (Professional)
;
;=============================================================

(defun LSEC:DrawText
       (pt txt height layer / ent)

  (if (and
        (listp pt)
        (= (length pt) 3)
        txt
        (> (strlen txt) 0)
        (numberp height)
        (> height 0)
        (LSEC:LayerExists layer)
      )

    (progn

      (setq ent

        (entmakex

          (list

            '(0 . "TEXT")

            '(100 . "AcDbEntity")

            (cons 8 layer)

            '(100 . "AcDbText")

            (cons 10 pt)

            (cons 40 height)

            (cons 1 txt)

            (cons 50 0.0)

            (cons 7 (getvar "TEXTSTYLE"))

            (cons 72 0)

            (cons 73 0)

          )

        )

      )

      (if ent

        ent

        (progn

          (LSEC:Print
            "Text creation failed.")

          nil

        )

      )

    )

    nil

  )

)

;=============================================================
;
; Part 5A.3 - Line Creator (Professional)
;
;=============================================================

(defun LSEC:DrawLine
       (p1 p2 layer / ent)

  (if (and
        (listp p1)
        (listp p2)
        (= (length p1) 3)
        (= (length p2) 3)
        (LSEC:LayerExists layer)
      )

    (progn

      (setq ent

        (entmakex

          (list

            '(0 . "LINE")

            '(100 . "AcDbEntity")

            (cons 8 layer)

            '(100 . "AcDbLine")

            (cons 10 p1)

            (cons 11 p2)

          )

        )

      )

      (if ent

        ent

        (progn

          (LSEC:Print
            "Line creation failed.")

          nil

        )

      )

    )

    nil

  )

)

;=============================================================
;
; Part 5A.4 - Circle Creator (Professional)
;
;=============================================================

(defun LSEC:DrawCircle
       (center radius layer / ent)

  (if (and
        (listp center)
        (= (length center) 3)
        (numberp radius)
        (> radius 0)
        (LSEC:LayerExists layer)
      )

    (progn

      (setq ent

        (entmakex

          (list

            '(0 . "CIRCLE")

            '(100 . "AcDbEntity")

            (cons 8 layer)

            '(100 . "AcDbCircle")

            (cons 10 center)

            (cons 40 radius)

          )

        )

      )

      (if ent

        ent

        (progn

          (LSEC:Print
            "Circle creation failed.")

          nil

        )

      )

    )

    nil

  )

)

;=============================================================
;
; Part 5B.1 - Draw EGL Profile
;
;=============================================================

(defun LSEC:DrawEGL (/ ent)

  (LSEC:CreateLayer
    "LSEC-EGL"
    1
  )

  (setq ent

    (LSEC:DrawPolyline
      (LSEC:GetEGLPoints)
      "LSEC-EGL"
    )

  )

  ent

)

;=============================================================
;
; Part 5B.2 - Draw BED Profile
;
;=============================================================

(defun LSEC:DrawBED (/ ent)

  (LSEC:CreateLayer
    "LSEC-BED"
    2
  )

  (setq ent

    (LSEC:DrawPolyline
      (LSEC:GetBEDPoints)
      "LSEC-BED"
    )

  )

  ent

)

;=============================================================
;
; Part 5B.3 - Draw FSL Profile
;
;=============================================================

(defun LSEC:DrawFSL (/ ent)

  (LSEC:CreateLayer
    "LSEC-FSL"
    4
  )

  (setq ent

    (LSEC:DrawPolyline
      (LSEC:GetFSLPoints)
      "LSEC-FSL"
    )

  )

  ent

)

;=============================================================
;
; Part 5B.4 - Draw BANK Profile
;
;=============================================================

(defun LSEC:DrawBANK (/ ent)

  (LSEC:CreateLayer
    "LSEC-BANK"
    3
  )

  (setq ent

    (LSEC:DrawPolyline
      (LSEC:GetBANKPoints)
      "LSEC-BANK"
    )

  )

  ent

)

;=============================================================
;
; Part 5B.5 - Draw All Profiles
;
;=============================================================

(defun LSEC:DrawProfiles (/)

  (LSEC:Print
    "Drawing profiles..."
  )

  (LSEC:DrawEGL)

  (LSEC:DrawBED)

  (LSEC:DrawFSL)

  (LSEC:DrawBANK)

  (LSEC:Print
    "Profiles completed."
  )

  T

)

;=============================================================
;
; Part 5C.1 - Grid Settings
;
;=============================================================

(defun LSEC:GetGridX (/ v)

  (setq v
    (LSEC:GetConfig "GRIDX")
  )

  (if v
    v
    20.0
  )

)

(defun LSEC:GetGridY (/ v)

  (setq v
    (LSEC:GetConfig "GRIDY")
  )

  (if v
    v
    1.0
  )

)

;=============================================================
;
; Part 5C.2 - Vertical Grid Lines
;
;=============================================================

(defun LSEC:DrawVerticalGrid
       (/ start
          end
          step
          ch
          x
          y1
          y2
          maxElev
          firstRec
          lastRec)

  ;; Make sure data exists

  (if (not *LSEC-DATA*)

    (progn

      (LSEC:Print
        "Vertical grid aborted: no Excel data loaded."
      )

      nil

    )

    (progn

      ;; Get first and last records

      (setq firstRec
            (car *LSEC-DATA*))

      (setq lastRec
            (car (reverse *LSEC-DATA*)))


      ;; Get first and last chainage

      (setq start
            (cdr
              (assoc "CHAINAGE" firstRec)
            )
      )

      (setq end
            (cdr
              (assoc "CHAINAGE" lastRec)
            )
      )


      ;; Grid spacing

      (setq step
            (LSEC:GetGridX))


      ;; Bottom of profile/grid

      (setq y1
            0.0)


      ;; Top of grid

      (setq maxElev
            (LSEC:FindMaxElevation))


      ;; Validate only numeric values

      (if (and
            (numberp start)
            (numberp end)
            (numberp step)
            (numberp maxElev)
            (> step 0.0)
          )

        (progn

          (setq y2
                (LSEC:ELtoY
                  (+ maxElev 5.0)
                )
          )


          ;; Draw vertical grid lines

          (setq ch start)

          (while (<= ch end)

            (setq x
                  (LSEC:CHtoX ch)
            )

            (if (and
                  (numberp x)
                  (numberp y1)
                  (numberp y2)
                )

              (LSEC:DrawLine
                (list x y1 0.0)
                (list x y2 0.0)
                "LSEC-GRID"
              )

            )

            (setq ch
                  (+ ch step)
            )

          )

          (LSEC:Print
            "Vertical grid completed."
          )

          T

        )

        (progn

          (LSEC:Print
            "Vertical grid aborted: invalid grid geometry."
          )

          nil

        )

      )

    )

  )

)

;=============================================================
;
; Part 5C.3 - Horizontal Grid Lines
;
;=============================================================

(defun LSEC:DrawHorizontalGrid
       (/ minElev maxElev elev step
          firstRec lastRec
          x1 x2 y)

  (setq minElev
        (LSEC:FindMinElevation))

  (setq maxElev
        (LSEC:FindMaxElevation))

  (setq step
        (LSEC:GetGridY))


  ;; Safety check

  (if (not
        (and
          (numberp minElev)
          (numberp maxElev)
          (numberp step)
          (> step 0.0)
          *LSEC-DATA*
        )
      )

    (progn

      (LSEC:Print
        "LSEC:DrawHorizontalGrid aborted: invalid grid data."
      )

      nil

    )

    (progn

      ;; Get actual first and last records

      (setq firstRec
            (car *LSEC-DATA*))

      (setq lastRec
            (car (reverse *LSEC-DATA*)))


      ;; Convert chainages to drawing X coordinates

      (setq x1
            (LSEC:CHtoX
              (cdr
                (assoc "CHAINAGE" firstRec)
              )
            )
      )

      (setq x2
            (LSEC:CHtoX
              (cdr
                (assoc "CHAINAGE" lastRec)
              )
            )
      )


      ;; Draw horizontal grid

      (setq elev minElev)

      (while (<= elev (+ maxElev 5.0))

        (setq y
              (LSEC:ELtoY elev))

        (if (and
              (numberp x1)
              (numberp x2)
              (numberp y)
            )

          (LSEC:DrawLine

            (list x1 y 0.0)

            (list x2 y 0.0)

            "LSEC-GRID"

          )

        )

        (setq elev
              (+ elev step))

      )


      (LSEC:Print
        "Horizontal grid completed."
      )

      T

    )

  )

)

;=============================================================
;
; Part 5C.4 - Grid Labels
;
;=============================================================

(defun LSEC:DrawGridLabels
       (/ step ch end x y elev minElev maxElev lastRec)

  ;-----------------------------------------------------------
  ; Get grid spacing
  ;-----------------------------------------------------------

  (setq step
        (LSEC:GetGridX))


  ;-----------------------------------------------------------
  ; Get actual last data record
  ;
  ; IMPORTANT:
  ; (last *LSEC-DATA*) returns a LIST containing the last
  ; record, so use CAR to get the record itself.
  ;-----------------------------------------------------------

  (setq lastRec
        (car (reverse *LSEC-DATA*)))


  (setq end
        (cdr
          (assoc
            "CHAINAGE"
            lastRec
          )
        )
  )


  ;-----------------------------------------------------------
  ; Chainage Labels
  ;-----------------------------------------------------------

  (setq ch 0.0)

  (while (and end (<= ch end))

    (setq x
          (LSEC:CHtoX ch))

    (LSEC:DrawText

      (list x -5.0 0.0)

      (rtos ch 2 0)

      (LSEC:GetConfig "TEXTHEIGHT")

      "LSEC-TEXT"

    )

    (setq ch
          (+ ch step))

  )


  ;-----------------------------------------------------------
  ; Elevation labels are drawn by LSEC:DrawProfileBoundaries
  ; so the numeric labels, ticks and left vertical axis share
  ; one reference-style geometry.
  ;-----------------------------------------------------------

  (LSEC:Print
    "Grid labels completed."
  )

  T

)

;=============================================================
;
; Part 5C.4A - Split-Aware Profile Boundaries
; and Elevation Axis
;
;=============================================================

(defun LSEC:DrawProfileBoundaries
       (/ firstRec
          lastRec
          start
          end
          x1
          x2
          datum
          maxElev
          y1
          y2
          tickLen
          elev
          step
          y)

  (if *LSEC-DATA*

    (progn

      (setq firstRec
        (car *LSEC-DATA*)
      )

      (setq lastRec
        (car (reverse *LSEC-DATA*))
      )

      (setq start
        (cdr (assoc "CHAINAGE" firstRec))
      )

      (setq end
        (cdr (assoc "CHAINAGE" lastRec))
      )

      (setq datum
        (LSEC:GetConfig "DATUM")
      )

      (setq maxElev
        (LSEC:FindMaxElevation)
      )

      (setq step
        (LSEC:GetGridY)
      )

      (setq tickLen 4.0)

      (if
        (and
          (numberp start)
          (numberp end)
          (numberp datum)
          (numberp maxElev)
          (numberp step)
          (> step 0.0)
        )

        (progn

          ;---------------------------------------------------
          ; IMPORTANT:
          ;
          ; Use ActiveCHtoX, not CHtoX.
          ;
          ; This makes the elevation axis and right boundary
          ; move with the active split frame.
          ;---------------------------------------------------

          (setq x1
            (LSEC:ActiveCHtoX start)
          )

          (setq x2
            (LSEC:ActiveCHtoX end)
          )

          (setq y1
            (LSEC:ELtoY datum)
          )

          (setq y2
            (LSEC:ELtoY
              (+ maxElev 5.0)
            )
          )

          ;---------------------------------------------------
          ; Beginning vertical line / elevation axis
          ;---------------------------------------------------

          (LSEC:DrawLine
            (list x1 y1 0.0)
            (list x1 y2 0.0)
            "LSEC-GRID"
          )

          ;---------------------------------------------------
          ; End vertical boundary line
          ;---------------------------------------------------

          (LSEC:DrawLine
            (list x2 y1 0.0)
            (list x2 y2 0.0)
            "LSEC-GRID"
          )

          ;---------------------------------------------------
          ; Elevation ticks and labels
          ;---------------------------------------------------

          (setq elev datum)

          (while (<= elev maxElev)

            (setq y
              (LSEC:ELtoY elev)
            )

            ; Tick line
            (LSEC:DrawLine
              (list (- x1 tickLen) y 0.0)
              (list x1 y 0.0)
              "LSEC-GRID"
            )

            ; RL label
            (LSEC:DrawText
              (list
                (- x1 18.0)
                (- y 1.0)
                0.0
              )
              (rtos elev 2 2)
              (LSEC:GetTextHeight)
              "LSEC-TEXT"
            )

            (setq elev
              (+ elev step)
            )

          )

          ;---------------------------------------------------
          ; Elevation axis title
          ;---------------------------------------------------

          (LSEC:DrawTextRotated
            (list
              (- x1 30.0)
              (/ (+ y1 y2) 2.0)
              0.0
            )
            "ELEVATION (m)"
            (LSEC:GetTextHeight)
            90.0
            "LSEC-TEXT"
          )

          (LSEC:Print
            "Split-aware profile boundaries and elevation axis completed."
          )

          T

        )

      )

    )

  )

)

;=============================================================
;
; Part 5J - Split-Aware Profile Legend
;
;=============================================================

(defun LSEC:DrawLegend
       (/ firstRec
          datum
          x1
          x
          y
          lineLen
          gap
          h
          items
          item)

  (if *LSEC-DATA*

    (progn

      (setq firstRec
        (car *LSEC-DATA*)
      )

      ;-------------------------------------------------------
      ; IMPORTANT:
      ;
      ; Get the active section start position.
      ;-------------------------------------------------------

      (setq x1
        (LSEC:ActiveCHtoX
          (cdr
            (assoc "CHAINAGE" firstRec)
          )
        )
      )

      (setq datum
        (LSEC:GetConfig "DATUM")
      )

      ; Put legend outside and to the left of this section's
      ; elevation axis.
      (setq x
        (- x1 125.0)
      )

      (setq y
        (+
          (LSEC:ELtoY datum)
          28.0
        )
      )

      (setq lineLen 28.0)
      (setq gap 6.0)
      (setq h (LSEC:GetTextHeight))

      (setq items
        (list
          (list "LSEC-EGL"  "GROUND LEVEL")
          (list "LSEC-BANK" "TOP BANK LEVEL")
          (list "LSEC-FSL"  "FULL SUPPLY LEVEL")
          (list "LSEC-BED"  "CANAL BED LEVEL")
        )
      )

      (foreach item items

        (LSEC:DrawLine
          (list x y 0.0)
          (list (+ x lineLen) y 0.0)
          (car item)
        )

        (LSEC:DrawText
          (list
            (+ x lineLen 5.0)
            (- y 1.25)
            0.0
          )
          (cadr item)
          h
          "LSEC-TEXT"
        )

        (setq y
          (- y gap)
        )

      )

      (LSEC:Print
        "Split-aware profile legend completed."
      )

      T

    )

  )

)

;=============================================================
;
; Part 5C.5 - Draw Complete Grid
;
;=============================================================

(defun LSEC:DrawGrid (/)

  (LSEC:Print
    "Drawing complete grid..."
  )

  (LSEC:DrawVerticalGrid)

  (LSEC:DrawHorizontalGrid)

  ;; Reference-style left/right profile boundaries and elevation axis.
  (LSEC:DrawProfileBoundaries)

  (LSEC:DrawGridLabels)

  (LSEC:Print
    "Grid completed."
  )

  T

)

;=============================================================
;
; Part 1C.x - Draw Rotated Text
;
;=============================================================

(defun LSEC:DrawTextRotated
       (pt txt height angle layer / oldLayer)

  ;; Save current layer
  (setq oldLayer (getvar "CLAYER"))

  ;; Change layer if supplied
  (if layer
    (setvar "CLAYER" layer)
  )

  ;; Create text
  (entmakex
    (list
      '(0 . "TEXT")
      (cons 8 layer)
      (cons 10 pt)
      (cons 40 height)
      (cons 1 txt)
      (cons 50 (* pi (/ angle 180.0)))
      (cons 7 (getvar "TEXTSTYLE"))
      (cons 72 1)          ; Center horizontally
      (cons 73 2)          ; Middle vertically
      (cons 11 pt)
    )
  )

  ;; Restore previous layer
  (if oldLayer
    (setvar "CLAYER" oldLayer)
  )

  T

)

;=============================================================
;
; Test Rotated Text
;
;=============================================================

(defun LSEC:TestRotatedText (/)

  (LSEC:CreateLayer
    "LSEC-TEXT"
    7
  )

  (LSEC:DrawTextRotated
    '(0 0 0)
    "90 DEG"
    3.5
    90
    "LSEC-TEXT"
  )

  (LSEC:DrawTextRotated
    '(20 0 0)
    "65 DEG"
    3.5
    65
    "LSEC-TEXT"
  )

  (LSEC:DrawTextRotated
    '(40 0 0)
    "0 DEG"
    3.5
    0
    "LSEC-TEXT"
  )

  (princ)

)

;=============================================================
;
; Part 5D.1 - Title Settings
;
;=============================================================

(defun LSEC:GetTitleHeight ()

  (or
    (LSEC:GetConfig "TITLEHEIGHT")
    4.0
  )

)


(defun LSEC:GetTextHeight (/ v)
  (setq v
    (LSEC:GetConfig "TEXTHEIGHT")
  )
  (if v
    v
    3.5
  )
)

;=============================================================
;
; Part 5D.2 - Draw Drawing Title
; Split-Aware Version
;
;=============================================================

(defun LSEC:DrawTitle
       (/ txt h firstRec startCH x y)

  ;-----------------------------------------------------------
  ; Title settings
  ;-----------------------------------------------------------

  (setq h
    (LSEC:GetTitleHeight)
  )

  (setq txt
    "CANAL LONGITUDINAL SECTION"
  )


  ;-----------------------------------------------------------
  ; Get active section start chainage.
  ;
  ; ActiveCHtoX will return:
  ;
  ; Normal drawing:
  ;   CHtoX(startCH)
  ;
  ; Split drawing:
  ;   FrameX + CHtoX(startCH - SectionStart)
  ;
  ; Therefore the title follows each split frame.
  ;-----------------------------------------------------------

  (if *LSEC-DATA*

    (setq firstRec
      (car *LSEC-DATA*)
    )

  )


  (if firstRec

    (setq startCH
      (cdr (assoc "CHAINAGE" firstRec))
    )

    (setq startCH 0.0)

  )


  ;-----------------------------------------------------------
  ; Split-aware X position.
  ;-----------------------------------------------------------

  (setq x
    (LSEC:ActiveCHtoX startCH)
  )


  ; Keep the existing vertical title position.
  (setq y 100.0)


  ;-----------------------------------------------------------
  ; Draw title
  ;-----------------------------------------------------------

  (LSEC:DrawText

    (list x y 0.0)

    txt

    h

    "LSEC-TEXT"

  )


  (LSEC:Print
    "Split-aware drawing title completed."
  )

  T

)

;=============================================================
;
; Part 5D.3 - Drawing Information Block
; Split-Aware Frame Version
;
;=============================================================

(defun LSEC:DrawInfoBlock
       (/ h datum hscale vexag x)

  ;-----------------------------------------------------------
  ; Get settings
  ;-----------------------------------------------------------

  (setq h
    (LSEC:GetTextHeight)
  )

  (setq datum
    (LSEC:GetConfig "DATUM")
  )

  (setq hscale
    (LSEC:GetConfig "HSCALE")
  )

  (setq vexag
    (LSEC:GetConfig "VEXAG")
  )


  ;-----------------------------------------------------------
  ; Determine X position
  ;
  ; Split drawing:
  ;   use active section frame X directly.
  ;
  ; Normal drawing:
  ;   use original X = 0.0.
  ;-----------------------------------------------------------

  (if (numberp *LSEC-SECTION-FRAME-X*)

    (setq x
      *LSEC-SECTION-FRAME-X*
    )

    (setq x 0.0)

  )


  ;-----------------------------------------------------------
  ; Project
  ;-----------------------------------------------------------

  (LSEC:DrawText
    (list x 90.0 0.0)
    "PROJECT : TEST PROJECT"
    h
    "LSEC-TEXT"
  )


  ;-----------------------------------------------------------
  ; Canal
  ;-----------------------------------------------------------

  (LSEC:DrawText
    (list x 85.0 0.0)
    "CANAL : MAIN CANAL"
    h
    "LSEC-TEXT"
  )


  ;-----------------------------------------------------------
  ; Horizontal Scale
  ;-----------------------------------------------------------

  (LSEC:DrawText
    (list x 80.0 0.0)
    (strcat
      "H-SCALE : 1:"
      (rtos hscale 2 0)
    )
    h
    "LSEC-TEXT"
  )


  ;-----------------------------------------------------------
  ; Vertical Exaggeration
  ;-----------------------------------------------------------

  (LSEC:DrawText
    (list x 75.0 0.0)
    (strcat
      "VERTICAL EXAGGERATION : "
      (rtos vexag 2 1)
    )
    h
    "LSEC-TEXT"
  )


  ;-----------------------------------------------------------
  ; Datum
  ;-----------------------------------------------------------

  (if datum

    (LSEC:DrawText
      (list x 70.0 0.0)
      (strcat
        "DATUM : "
        (rtos datum 2 3)
      )
      h
      "LSEC-TEXT"
    )

  )


  (LSEC:Print
    "Split-aware project information block completed."
  )

  T

)


;=============================================================
;
; Part 5D.4 - Dynamic Sheet Border and Layout Frame
;
;=============================================================

(defun LSEC:DrawSheetFrame (/ start end maxElev datum x1 x2 y1 y2)

  (setq start
    (cdr
      (assoc "CHAINAGE"
        (car *LSEC-DATA*)
      )
    )
  )

  (setq end
    (cdr
      (assoc "CHAINAGE"
        (last *LSEC-DATA*)
      )
    )
  )


  (setq maxElev
    (LSEC:FindMaxElevation)
  )


  (setq datum
    (LSEC:GetConfig "DATUM")
  )


  (if (and start end maxElev datum)

    (progn

      ; Convert profile limits to drawing coordinates

      (setq x1
        (- (LSEC:CHtoX start) 50.0)
      )


      (setq x2
        (+ (LSEC:CHtoX end) 100.0)
      )


      (setq y1
        (- (LSEC:ELtoY datum) 30.0)
      )


      (setq y2
        (+ (LSEC:ELtoY (+ maxElev 5.0)) 40.0)
      )


      ; Bottom

      (LSEC:DrawLine
        (list x1 y1 0.0)
        (list x2 y1 0.0)
        "LSEC-TEXT"
      )


      ; Top

      (LSEC:DrawLine
        (list x1 y2 0.0)
        (list x2 y2 0.0)
        "LSEC-TEXT"
      )


      ; Left

      (LSEC:DrawLine
        (list x1 y1 0.0)
        (list x1 y2 0.0)
        "LSEC-TEXT"
      )


      ; Right

      (LSEC:DrawLine
        (list x2 y1 0.0)
        (list x2 y2 0.0)
        "LSEC-TEXT"
      )


      (LSEC:Print
        "Dynamic sheet frame completed."
      )


      T

    )


    (progn

      (LSEC:Print
        "Sheet frame aborted: missing profile data."
      )

      nil

    )

  )

)

;=============================================================
;
; Part 5D.5 - Title Block Layout
; Split-Aware Version
;
;=============================================================

(defun LSEC:DrawTitleBlock
       (/ start
          end
          datum
          hscale
          vexag
          maxElev
          minElev
          firstRec
          lastRec
          x
          y
          h)

  ;-----------------------------------------------------------
  ; Validate Excel data
  ;-----------------------------------------------------------

  (if (not *LSEC-DATA*)

    (progn

      (LSEC:Print
        "Title block aborted: no Excel data loaded."
      )

      nil

    )

    (progn

      ;-------------------------------------------------------
      ; Get first and last records
      ;-------------------------------------------------------

      (setq firstRec
        (car *LSEC-DATA*)
      )

      (setq lastRec
        (car (reverse *LSEC-DATA*))
      )


      ;-------------------------------------------------------
      ; Get chainage limits
      ;-------------------------------------------------------

      (setq start
        (cdr (assoc "CHAINAGE" firstRec))
      )

      (setq end
        (cdr (assoc "CHAINAGE" lastRec))
      )


      ;-------------------------------------------------------
      ; Get elevation limits
      ;-------------------------------------------------------

      (setq minElev
        (LSEC:FindMinElevation)
      )

      (setq maxElev
        (LSEC:FindMaxElevation)
      )


      ;-------------------------------------------------------
      ; Get datum
      ;-------------------------------------------------------

      (setq datum
        (LSEC:GetConfig "DATUM")
      )

      (if (not (numberp datum))

        (if (numberp minElev)

          (progn

            (setq datum
              (fix minElev)
            )

            (LSEC:SetConfig
              "DATUM"
              datum
            )

          )

        )

      )


      ;-------------------------------------------------------
      ; Other settings
      ;-------------------------------------------------------

      (setq h
        (LSEC:GetTextHeight)
      )

      (setq hscale
        (LSEC:GetConfig "HSCALE")
      )

      (if (not hscale)
        (setq hscale 1000.0)
      )

      (setq vexag
        (LSEC:GetConfig "VEXAG")
      )

      (if (not vexag)
        (setq vexag 10.0)
      )


      ;-------------------------------------------------------
      ; Validate geometry
      ;-------------------------------------------------------

      (if (and
            (numberp start)
            (numberp end)
            (numberp datum)
            (numberp minElev)
            (numberp maxElev)
          )

        (progn


          ;---------------------------------------------------
          ; IMPORTANT - SPLIT-AWARE X POSITION
          ;
          ; Normal drawing:
          ;   ActiveCHtoX(start) = normal CHtoX(start)
          ;
          ; Split drawing:
          ;   ActiveCHtoX(start) =
          ;   Frame X + CHtoX(start - Section Start)
          ;
          ; Therefore the complete title block follows
          ; the respective split drawing.
          ;---------------------------------------------------

          (setq x
            (LSEC:ActiveCHtoX start)
          )


          ;---------------------------------------------------
          ; Title block vertical position
          ;---------------------------------------------------

          (setq y
            (+
              (LSEC:ELtoY
                (+ maxElev 5.0)
              )
              20.0
            )
          )


          ;---------------------------------------------------
          ; Main title
          ;---------------------------------------------------

          (LSEC:DrawText
            (list x y 0.0)
            "CANAL LONGITUDINAL SECTION"
            (LSEC:GetTitleHeight)
            "LSEC-TEXT"
          )


          ;---------------------------------------------------
          ; Project information
          ;---------------------------------------------------

          (setq y
            (- y 10.0)
          )

          (LSEC:DrawText
            (list x y 0.0)
            "PROJECT : TEST PROJECT"
            h
            "LSEC-TEXT"
          )


          (setq y
            (- y 5.0)
          )

          (LSEC:DrawText
            (list x y 0.0)
            "DRAWING : LONGITUDINAL PROFILE"
            h
            "LSEC-TEXT"
          )


          (setq y
            (- y 5.0)
          )

          (LSEC:DrawText
            (list x y 0.0)
            (strcat
              "H-SCALE : 1:"
              (rtos hscale 2 0)
            )
            h
            "LSEC-TEXT"
          )


          (setq y
            (- y 5.0)
          )

          (LSEC:DrawText
            (list x y 0.0)
            (strcat
              "VERTICAL EXAGGERATION : "
              (rtos vexag 2 1)
            )
            h
            "LSEC-TEXT"
          )


          (setq y
            (- y 5.0)
          )

          (LSEC:DrawText
            (list x y 0.0)
            (strcat
              "DATUM : "
              (rtos datum 2 3)
            )
            h
            "LSEC-TEXT"
          )


          (LSEC:Print
            "Split-aware title block completed."
          )

          T

        )

        (progn

          (LSEC:Print
            "Title block aborted: missing profile data."
          )

          nil

        )

      )

    )

  )

)


;=============================================================
;
; Part 5D.6 - Draw Structures
;
;=============================================================

;=============================================================
;
; Part 5D.6 - Draw Structures (Reference-Style Placement)
;
;=============================================================

(defun LSEC:DrawStructures
       (/ rec ch structName x egl bank fsl bed topElev yBase flagHeight textOffset)

  (LSEC:CreateLayer "LSEC-STRUCT" 1)

  ;; The structure name must clear the highest profile at its station.
  (setq flagHeight 15.0)
  (setq textOffset 3.0)

  (foreach rec *LSEC-DATA*

    (setq structName (cdr (assoc "STRUCT" rec)))

    (if (and structName (/= structName ""))
      (progn

        (setq ch   (cdr (assoc "CHAINAGE" rec)))
        (setq egl  (cdr (assoc "EGL" rec)))
        (setq bank (cdr (assoc "BANK" rec)))
        (setq fsl  (cdr (assoc "FSL" rec)))
        (setq bed  (cdr (assoc "BED" rec)))

        ;; Use the highest valid profile elevation as the flag base.
        (setq topElev nil)
        (foreach elev (list egl bank fsl bed)
          (if (and (numberp elev)
                   (or (not topElev) (> elev topElev)))
            (setq topElev elev)
          )
        )

        (setq x (LSEC:CHtoX ch))
        (setq yBase (if topElev (LSEC:ELtoY topElev) nil))

        (if (and (numberp x) (numberp yBase))
          (progn

            (LSEC:DrawLine
              (list x yBase 0.0)
              (list x (+ yBase flagHeight) 0.0)
              "LSEC-STRUCT"
            )

            ;; Offset right and upward so even the last structure label
            ;; remains above the profile instead of falling below it.
            (LSEC:DrawTextRotated
              (list (+ x 2.0)
                    (+ yBase flagHeight textOffset)
                    0.0)
              structName
              (LSEC:GetTextHeight)
              65.0
              "LSEC-STRUCT"
            )
          )
        )
      )
    )
  )

  (LSEC:Print "Structures completed.")
  T

)

;=============================================================
;
; Test Part 5D.6
;
;=============================================================

(defun LSEC:TestStructures (/)

  (if (null *LSEC-DATA*)

    (progn

      (LSEC:Print
        "No Excel data loaded."
      )

    )

    (LSEC:DrawStructures)

  )

)

;=============================================================
;
; L-SECV6 Professional V6
; Part 5E.1 - Table Configuration
;
; Purpose:
;   Store fixed table presentation configuration.
;
; This part does NOT calculate geometry.
; This part does NOT draw the table.
;
;=============================================================


;-------------------------------------------------------------
; Initialize Table Configuration
;-------------------------------------------------------------

(defun LSEC:InitTableConfig (/)

  ;-----------------------------------------------------------
  ; Table Row Height
  ;
  ; Seven rows:
  ; CUT
  ; FILL
  ; TOP BANK
  ; FSL
  ; BED
  ; GROUND
  ; DISTANCE
  ;
  ;-----------------------------------------------------------

  (setq *LSEC-CONFIG*
    (subst
      (cons "TABLEROWHEIGHT" 30.0)
      (assoc "TABLEROWHEIGHT" *LSEC-CONFIG*)
      *LSEC-CONFIG*
    )
  )

  (if (not (assoc "TABLEROWHEIGHT" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (cons
        '("TABLEROWHEIGHT" . 30.0)
        *LSEC-CONFIG*
      )
    )
  )


  ;-----------------------------------------------------------
  ; Table Label Column Width
  ;
  ; This is the area for:
  ;
  ; CUT DEPTH
  ; FILL DEPTH
  ; TOP BANK LEVEL
  ; FULL SUPPLY LEVEL
  ; CANAL BED LEVEL
  ; GROUND LEVEL
  ; DISTANCE IN METRES
  ;
  ;-----------------------------------------------------------

  (setq *LSEC-CONFIG*
    (subst
      (cons "TABLELABELWIDTH" 60.0)
      (assoc "TABLELABELWIDTH" *LSEC-CONFIG*)
      *LSEC-CONFIG*
    )
  )

  (if (not (assoc "TABLELABELWIDTH" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (cons
        '("TABLELABELWIDTH" . 60.0)
        *LSEC-CONFIG*
      )
    )
  )


  ;-----------------------------------------------------------
  ; Gap Between Profile And Table
  ;-----------------------------------------------------------

  (setq *LSEC-CONFIG*
    (subst
      (cons "TABLEGAP" 10.0)
      (assoc "TABLEGAP" *LSEC-CONFIG*)
      *LSEC-CONFIG*
    )
  )

  (if (not (assoc "TABLEGAP" *LSEC-CONFIG*))
    (setq *LSEC-CONFIG*
      (cons
        '("TABLEGAP" . 10.0)
        *LSEC-CONFIG*
      )
    )
  )


  ;-----------------------------------------------------------
  ; Report
  ;-----------------------------------------------------------

;-----------------------------------------------------------
; Report
;-----------------------------------------------------------

  (if *LSEC-DEBUG*
    (LSEC:Print
      "Part 5E.1 Table Configuration Initialized."
    )
  )

  T

)


;=============================================================
;
; Part 5E.1 Test
;
;=============================================================

(defun LSEC:TestTableConfig
       (/ rowHeight labelWidth gap)

  ; Initialize configuration

  (LSEC:InitTableConfig)


  ; Read values

  (setq rowHeight
    (cdr
      (assoc
        "TABLEROWHEIGHT"
        *LSEC-CONFIG*
      )
    )
  )


  (setq labelWidth
    (cdr
      (assoc
        "TABLELABELWIDTH"
        *LSEC-CONFIG*
      )
    )
  )


  (setq gap
    (cdr
      (assoc
        "TABLEGAP"
        *LSEC-CONFIG*
      )
    )
  )


  ; Report values

  (LSEC:Print
    (strcat
      "TABLEROWHEIGHT = "
      (rtos rowHeight 2 2)
    )
  )


  (LSEC:Print
    (strcat
      "TABLELABELWIDTH = "
      (rtos labelWidth 2 2)
    )
  )


  (LSEC:Print
    (strcat
      "TABLEGAP = "
      (rtos gap 2 2)
    )
  )


  (LSEC:Print
    "Part 5E.1 Test Completed."
  )

  T

)


;=============================================================
;
; Part 5E.1A - Ensure Table Layer
;
;=============================================================

(defun LSEC:EnsureTableLayer (/)

  (if (not (LSEC:LayerExists "LSEC-TABLE"))

    (progn

      (LSEC:CreateLayer
        "LSEC-TABLE"
        7
      )

    )

  )

  T

)

;=============================================================
;
; Part 5E.2 - Table Geometry
;
; Single source of truth for overall table geometry.
;
;=============================================================


;-------------------------------------------------------------
; 5E.2.1 - Get Number of Stations
;-------------------------------------------------------------

(defun LSEC:GetStationCount (/)

  (if *LSEC-DATA*
    (length *LSEC-DATA*)
    0
  )

)

;-------------------------------------------------------------
; 5E.2.2 - Get Label Column Width
;-------------------------------------------------------------

(defun LSEC:GetLabelColumnWidth ()

  (cdr
    (assoc
      "TABLELABELWIDTH"
      *LSEC-CONFIG*
    )
  )

)


;-------------------------------------------------------------
; 5E.2.3 - Get Number of Table Rows
;-------------------------------------------------------------

(defun LSEC:GetRowCount (/)

  7

)


;-------------------------------------------------------------
; Compatibility name
;-------------------------------------------------------------

(defun LSEC:GetTableRowCount (/)

  (LSEC:GetRowCount)

)


;-------------------------------------------------------------
; 5E.2.3 - Get Table Width
;
; Total table width =
; Label column width + station-grid width
;
;-------------------------------------------------------------

(defun LSEC:GetTableWidth
       (/ stations startStation endStation
          startX endX labelWidth)

  (setq stations
        (LSEC:GetStationList))

  (if stations

    (progn

      (setq startStation
            (car stations))

      (setq endStation
            (nth
              (1- (length stations))
              stations
            )
      )

      (setq startX
            (LSEC:GetStationX startStation))

      (setq endX
            (LSEC:GetStationX endStation))

      (setq labelWidth
            (LSEC:GetLabelColumnWidth))

      (+

        labelWidth

        (- endX startX)

      )

    )

    0.0

  )

)

;-------------------------------------------------------------
; 5E.2.5 - Get Table Height
;-------------------------------------------------------------

(defun LSEC:GetTableHeight
       (/ rowCount rowHeight)

  (setq rowCount
        (LSEC:GetRowCount))

  (setq rowHeight
        (LSEC:GetRowHeight))


  (* rowCount rowHeight)

)

;=============================================================
;
; Part 5E.2.4 - Get Grid Base Point
;
; Split-aware version
;
;=============================================================

(defun LSEC:GetGridBasePoint
       (/ firstRec startCh datum x y)

  (if *LSEC-DATA*

    (progn

      ; Get first station of current active section
      (setq firstRec
        (car *LSEC-DATA*)
      )

      (setq startCh
        (cdr
          (assoc "CHAINAGE" firstRec)
        )
      )

      ; Get datum
      (setq datum
        (LSEC:GetConfig "DATUM")
      )

      ; IMPORTANT:
      ; Use split-aware chainage conversion.
      (if (numberp startCh)
        (setq x
          (LSEC:ActiveCHtoX startCh)
        )
        (setq x 0.0)
      )

      ; Datum Y coordinate
      (if (numberp datum)
        (setq y
          (LSEC:ELtoY datum)
        )
        (setq y 0.0)
      )

      (list x y 0.0)

    )

    (list 0.0 0.0 0.0)

  )

)


;=============================================================
;
; Part 5E.2.7 - Get Table Base Point
;
; Purpose:
;   Return the TOP-LEFT corner of the complete table.
;
; Split-aware:
;   The first station of the CURRENT active section is the
;   table/profile alignment reference.
;
;=============================================================

(defun LSEC:GetTableBasePoint
       (/ gridBase
          labelWidth
          gap
          x
          y)

  ;-----------------------------------------------------------
  ; Make sure table configuration exists
  ;-----------------------------------------------------------

  (LSEC:InitTableConfig)


  ;-----------------------------------------------------------
  ; Get profile/grid reference point
  ;
  ; GetGridBasePoint uses the first record of the currently
  ; active *LSEC-DATA* section.
  ;-----------------------------------------------------------

  (setq gridBase
        (LSEC:GetGridBasePoint)
  )


  ;-----------------------------------------------------------
  ; Get label column width
  ;-----------------------------------------------------------

  (setq labelWidth
        (LSEC:GetLabelColumnWidth)
  )


  ;-----------------------------------------------------------
  ; Get profile-to-table gap
  ;-----------------------------------------------------------

  (setq gap
        (LSEC:GetConfig "TABLEGAP")
  )


  ;-----------------------------------------------------------
  ; X coordinate
  ;
  ; gridBase X is the first station of the CURRENT section.
  ;
  ; The table extends labelWidth units to the left of that
  ; first station.
  ;-----------------------------------------------------------

  (setq x
        (- (car gridBase)
           labelWidth
        )
  )


  ;-----------------------------------------------------------
  ; Y coordinate
  ;
  ; Move table down from grid base by TABLEGAP.
  ;-----------------------------------------------------------

  (setq y
        (- (cadr gridBase)
           gap
        )
  )


  ;-----------------------------------------------------------
  ; Return table top-left point
  ;-----------------------------------------------------------

  (list x y 0.0)

)


;=============================================================
;
; Part 5E.3 - Row Geometry
;
;=============================================================


;-------------------------------------------------------------
; 5E.3.1 - Get Row Height
;-------------------------------------------------------------

(defun LSEC:GetRowHeight (/ h)

  (setq h
        (cdr
          (assoc
            "TABLEROWHEIGHT"
            *LSEC-CONFIG*
          )
        )
  )


  (if h
    h
    8.0
  )

)


;-------------------------------------------------------------
; 5E.3.2 - Get Row Top
;
; Row 0 starts at table top.
;
;-------------------------------------------------------------

(defun LSEC:GetRowTop
       (rowIndex / baseY rowHeight)

  (setq baseY
        (cadr
          (LSEC:GetTableBasePoint)
        ))


  (setq rowHeight
        (LSEC:GetRowHeight))


  (- baseY
     (* rowIndex rowHeight)
  )

)


;-------------------------------------------------------------
; 5E.3.3 - Get Row Center
;-------------------------------------------------------------

(defun LSEC:GetRowCenter
       (rowIndex / rowTop rowHeight)

  (setq rowTop
        (LSEC:GetRowTop rowIndex))


  (setq rowHeight
        (LSEC:GetRowHeight))


  (- rowTop
     (/ rowHeight 2.0)
  )

)


;-------------------------------------------------------------
; 5E.3.4 - Get Row Bottom
;-------------------------------------------------------------

(defun LSEC:GetRowBottom
       (rowIndex / rowTop rowHeight)

  (setq rowTop
        (LSEC:GetRowTop rowIndex))


  (setq rowHeight
        (LSEC:GetRowHeight))


  (- rowTop rowHeight)

)


;-------------------------------------------------------------
; 5E.3.5 - Get Row Y
;
; Compatibility function.
; Returns row center Y.
;
;-------------------------------------------------------------

(defun LSEC:GetRowY
       (rowIndex)

  (LSEC:GetRowCenter rowIndex)

)


;-------------------------------------------------------------
; 5E.3.6 - Test Row Geometry
;-------------------------------------------------------------

(defun LSEC:TestRowGeometry
       (/ i count)

  (LSEC:InitTableConfig)


  (setq count
        (LSEC:GetRowCount))


  (setq i 0)


  (LSEC:Print
    "----- Part 5E.3 Row Geometry -----"
  )


  (while (< i count)

    (LSEC:Print
      (strcat
        "Row "
        (itoa i)
        "  Top="
        (rtos
          (LSEC:GetRowTop i)
          2
          2
        )
        "  Center="
        (rtos
          (LSEC:GetRowCenter i)
          2
          2
        )
        "  Bottom="
        (rtos
          (LSEC:GetRowBottom i)
          2
          2
        )
      )
    )


    (setq i
          (1+ i))

  )


  (LSEC:Print
    "Part 5E.3 Test Completed."
  )


  T

)

;=============================================================
;
; Part 5E.4 - Station Geometry
;
; Purpose:
;   Single source of truth for all station X coordinates.
;
; This part does NOT draw anything.
; This part does NOT calculate row Y coordinates.
;
;=============================================================


;-------------------------------------------------------------
; 5E.4.1 - Get Station List
;
; Returns the chainage list from the existing L-SECV6
; data structure.
;
;-------------------------------------------------------------

(defun LSEC:GetStationList
       (/ result item)

  (setq result nil)

  (if *LSEC-DATA*

    (foreach item *LSEC-DATA*

      (if (assoc "CHAINAGE" item)

        (setq result
          (append
            result
            (list
              (cdr
                (assoc "CHAINAGE" item)
              )
            )
          )
        )

      )

    )

  )

  result

)


;-------------------------------------------------------------
; 5E.4.2 - Get Station X
;
; Split-aware station X coordinate.
;
;-------------------------------------------------------------

(defun LSEC:GetStationX
       (chainage)

  (LSEC:ActiveCHtoX chainage)

)


;-------------------------------------------------------------
; 5E.4.3 - Get Station Center
;
; Returns the X coordinate of the station column center.
;
;-------------------------------------------------------------

(defun LSEC:GetStationCenter
       (chainage)

  (LSEC:GetStationX chainage)

)


;-------------------------------------------------------------
; 5E.4.4 - Get Column Width
;
; Uses consecutive station X coordinates.
;
;-------------------------------------------------------------

(defun LSEC:GetColumnWidth
       (/ stations first second)

  (setq stations
        (LSEC:GetStationList))


  (if (> (length stations) 1)

    (progn

      (setq first
            (car stations))

      (setq second
            (cadr stations))


      (- (LSEC:GetStationX second)
         (LSEC:GetStationX first)
      )

    )

    0.0

  )

)


;-------------------------------------------------------------
; 5E.4.5 - Test Station Geometry
;-------------------------------------------------------------

(defun LSEC:TestStationGeometry
       (/ stations first lastStation)

  (setq stations
        (LSEC:GetStationList))


  (LSEC:Print
    "----- Part 5E.4 Station Geometry -----"
  )


  (LSEC:Print
    (strcat
      "Station Count = "
      (itoa (length stations))
    )
  )


  (if stations

    (progn

      (setq first
            (car stations))

      ; Get last station without using LAST
      (setq lastStation
            (car
              (reverse stations)
            ))


      (LSEC:Print
        (strcat
          "First Station = "
          (rtos first 2 2)
          "  X = "
          (rtos
            (LSEC:GetStationX first)
            2
            3
          )
        )
      )


      (LSEC:Print
        (strcat
          "Last Station = "
          (rtos lastStation 2 2)
          "  X = "
          (rtos
            (LSEC:GetStationX lastStation)
            2
            3
          )
        )
      )


      (LSEC:Print
        (strcat
          "Column Width = "
          (rtos
            (LSEC:GetColumnWidth)
            2
            3
          )
        )
      )

    )

  )


  (LSEC:Print
    "Part 5E.4 Test Completed."
  )


  T

)



;=============================================================
;
; Part 5E.5 - Cell Geometry
;
; Purpose:
;   Combine station X geometry and row Y geometry.
;
; This part does NOT draw anything.
;
;=============================================================


;-------------------------------------------------------------
; 5E.5.1 - Get Cell Center
;
; rowIndex = table row
; chainage = station
;
; Returns:
;   (X Y 0.0)
;
;-------------------------------------------------------------

(defun LSEC:GetCellCenter
       (rowIndex chainage / x y)

  (setq x
        (LSEC:GetStationCenter chainage))


  (setq y
        (LSEC:GetRowCenter rowIndex))


  (list
    x
    y
    0.0
  )

)


;-------------------------------------------------------------
; 5E.5.2 - Get Cell Top Left
;-------------------------------------------------------------

(defun LSEC:GetCellTopLeft
       (rowIndex chainage / x y)

  (setq x
        (LSEC:GetStationX chainage))


  (setq y
        (LSEC:GetRowTop rowIndex))


  (list
    x
    y
    0.0
  )

)


;-------------------------------------------------------------
; 5E.5.3 - Get Cell Bottom Left
;-------------------------------------------------------------

(defun LSEC:GetCellBottomLeft
       (rowIndex chainage / x y)

  (setq x
        (LSEC:GetStationX chainage))


  (setq y
        (LSEC:GetRowBottom rowIndex))


  (list
    x
    y
    0.0
  )

)


;-------------------------------------------------------------
; 5E.5.4 - Test Cell Geometry
;-------------------------------------------------------------

(defun LSEC:TestCellGeometry
       (/ stations first)

  (setq stations
        (LSEC:GetStationList))


  (LSEC:Print
    "----- Part 5E.5 Cell Geometry -----"
  )


  (if stations

    (progn

      (setq first
            (car stations))


      (LSEC:Print
        (strcat
          "Cell Center Row 0 @ CH "
          (rtos first 2 2)
          " = "
          (vl-princ-to-string
            (LSEC:GetCellCenter
              0
              first
            )
          )
        )
      )


      (LSEC:Print
        (strcat
          "Cell Top Left Row 0 @ CH "
          (rtos first 2 2)
          " = "
          (vl-princ-to-string
            (LSEC:GetCellTopLeft
              0
              first
            )
          )
        )
      )


      (LSEC:Print
        (strcat
          "Cell Bottom Left Row 0 @ CH "
          (rtos first 2 2)
          " = "
          (vl-princ-to-string
            (LSEC:GetCellBottomLeft
              0
              first
            )
          )
        )
      )

    )

  )


  (LSEC:Print
    "Part 5E.5 Test Completed."
  )


  T

)

;=============================================================
;
; L-SECV6 Professional V6
; Part 5E.6 - Draw Clean Table Frame
;
; Purpose:
;   Draw the professional lower table frame.
;
; The table contains:
;   - Outer boundary
;   - Horizontal row separator lines
;   - Label column separator
;
; IMPORTANT:
;   No vertical station lines are drawn inside the table.
;
; Split-aware:
;   The label separator is calculated from the current table
;   base point, so it moves together with the active section.
;
;=============================================================


;-------------------------------------------------------------
; Part 5E.6.1 - Draw One Table Line
;-------------------------------------------------------------

(defun LSEC:DrawTableFrameLine
       (pt1 pt2)

  (entmakex
    (list
      '(0 . "LINE")
      '(100 . "AcDbEntity")

      (cons 8 "LSEC-TABLE")
      (cons 62 8)

      (cons 10 pt1)
      (cons 11 pt2)
    )
  )

)


;-------------------------------------------------------------
; Part 5E.6.2 - Draw Table Frame
;-------------------------------------------------------------

(defun LSEC:DrawTableFrame
       (/ basePoint
          tableWidth
          tableHeight
          rowCount
          rowHeight
          labelWidth
          leftX
          rightX
          topY
          bottomY
          labelRightX
          rowIndex
          y)

  (LSEC:Print
    "Drawing Part 5E.6 Clean Table Frame..."
  )


  ;-----------------------------------------------------------
  ; Initialize table configuration
  ;-----------------------------------------------------------

  (LSEC:InitTableConfig)

  (LSEC:EnsureTableLayer)


  ;-----------------------------------------------------------
  ; Get table geometry
  ;-----------------------------------------------------------

  (setq basePoint
        (LSEC:GetTableBasePoint))

  (setq tableWidth
        (LSEC:GetTableWidth))

  (setq tableHeight
        (LSEC:GetTableHeight))

  (setq rowCount
        (LSEC:GetRowCount))

  (setq rowHeight
        (LSEC:GetRowHeight))

  (setq labelWidth
        (LSEC:GetLabelColumnWidth))


  ;-----------------------------------------------------------
  ; Calculate table boundaries
  ;
  ; Base point is the table top-left corner.
  ;-----------------------------------------------------------

  (setq leftX
        (car basePoint))

  (setq topY
        (cadr basePoint))

  (setq rightX
        (+ leftX tableWidth))

  (setq bottomY
        (- topY tableHeight))


  ;-----------------------------------------------------------
  ; Label column right boundary
  ;
  ; IMPORTANT FOR SPLIT DRAWING:
  ;
  ; Do NOT use:
  ;
  ;   (LSEC:GetStationX 0.0)
  ;
  ; because chainage 0.0 belongs only to the first section.
  ;
  ; The separator must always be positioned relative to the
  ; current table base point.
  ;-----------------------------------------------------------

  (setq labelRightX
        (+ leftX labelWidth))


  ;-----------------------------------------------------------
  ; Draw outer table boundary
  ;-----------------------------------------------------------

  ; Top
  (LSEC:DrawTableFrameLine
    (list leftX topY 0.0)
    (list rightX topY 0.0)
  )

  ; Bottom
  (LSEC:DrawTableFrameLine
    (list leftX bottomY 0.0)
    (list rightX bottomY 0.0)
  )

  ; Left
  (LSEC:DrawTableFrameLine
    (list leftX topY 0.0)
    (list leftX bottomY 0.0)
  )

  ; Right
  (LSEC:DrawTableFrameLine
    (list rightX topY 0.0)
    (list rightX bottomY 0.0)
  )


  ;-----------------------------------------------------------
  ; Draw label column separator
  ;-----------------------------------------------------------

  (LSEC:DrawTableFrameLine
    (list labelRightX topY 0.0)
    (list labelRightX bottomY 0.0)
  )


  ;-----------------------------------------------------------
  ; Draw horizontal row separators
  ;
  ; There are 7 rows.
  ; Internal separators = rowCount - 1
  ;-----------------------------------------------------------

  (setq rowIndex 1)

  (while (< rowIndex rowCount)

    (setq y
          (- topY
             (* rowIndex rowHeight)))

    (LSEC:DrawTableFrameLine
      (list leftX y 0.0)
      (list rightX y 0.0)
    )

    (setq rowIndex
          (1+ rowIndex))

  )


  ;-----------------------------------------------------------
  ; No vertical station lines are drawn here.
  ;
  ; Station/data text remains controlled by the Part 5E
  ; station geometry engine.
  ;-----------------------------------------------------------


  (LSEC:Print
    "Part 5E.6 Clean Split-Aware Table Frame Completed."
  )

  T

)

;=============================================================
;
; L-SECV6 Professional V6
; Part 5F - Row Titles
;
; Purpose:
;   Draw row titles inside the left label column.
;
; Uses Part 5E as the single source of truth for:
;   - Table base point
;   - Row centers
;   - Label column width
;
;=============================================================


;-------------------------------------------------------------
; Part 5F.1 - Get Row Title List
;-------------------------------------------------------------

(defun LSEC:GetRowTitleList ()

  (list
    "CUT DEPTH (m)"
    "FILL DEPTH (m)"
    "TOP BANK LEVEL (m)"
    "FULL SUPPLY LEVEL (m)"
    "CANAL BED LEVEL (m)"
    "GROUND LEVEL (m)"
    "DISTANCE IN MTRS."
  )

)


;-------------------------------------------------------------
; Part 5F.2 - Get Row Title
;
; rowIndex = 0 to 6
;-------------------------------------------------------------

(defun LSEC:GetRowTitle
       (rowIndex / titles)

  (setq titles
        (LSEC:GetRowTitleList))

  (if (and
        (>= rowIndex 0)
        (< rowIndex (length titles))
      )

    (nth rowIndex titles)

    nil

  )

)


;-------------------------------------------------------------
; Part 5F.3 - Get Row Title X
;
; Returns horizontal center of label column.
;
; Table base point is the top-left corner of the
; complete table, including the label column.
;-------------------------------------------------------------

(defun LSEC:GetRowTitleX
       (/ base labelWidth)

  (LSEC:InitTableConfig)

  (setq base
        (LSEC:GetTableBasePoint))

  (setq labelWidth
        (LSEC:GetLabelColumnWidth))

  (+
    (car base)
    (/ labelWidth 2.0)
  )

)


;-------------------------------------------------------------
; Part 5F.4 - Get Row Title Point
;
; Returns insertion point for a row title.
;
; X = center of label column
; Y = row center from Part 5E
;-------------------------------------------------------------

(defun LSEC:GetRowTitlePoint
       (rowIndex)

  (list
    (LSEC:GetRowTitleX)
    (LSEC:GetRowCenter rowIndex)
    0.0
  )

)


;-------------------------------------------------------------
; Part 5F.5 - Draw One Row Title
;-------------------------------------------------------------

(defun LSEC:DrawRowTitle
       (rowIndex / title pt textHeight)

  (setq title
        (LSEC:GetRowTitle rowIndex))

  (setq pt
        (LSEC:GetRowTitlePoint rowIndex))

  (setq textHeight
        (LSEC:GetConfig "TEXTHEIGHT"))

  (if (null textHeight)

    (setq textHeight 3.5)

  )

  (if title

    (progn

      (entmakex
        (list
          '(0 . "TEXT")
          '(100 . "AcDbEntity")
          '(100 . "AcDbText")
          (cons 8 "LSEC-TABLE")
          (cons 10 pt)
          (cons 40 textHeight)
          (cons 1 title)
          (cons 7 (getvar "TEXTSTYLE"))
          (cons 72 1)
          (cons 73 2)
          (cons 11 pt)
        )
      )

    )

  )

  T

)


;-------------------------------------------------------------
; Part 5F.6 - Draw All Row Titles
;-------------------------------------------------------------

(defun LSEC:DrawRowTitles
       (/ rowIndex rowCount)

  (LSEC:Print
    "Drawing Part 5F Row Titles..."
  )

  (LSEC:InitTableConfig)

  (LSEC:EnsureTableLayer)

  (setq rowIndex 0)

  (setq rowCount
        (LSEC:GetRowCount))

  (while (< rowIndex rowCount)

    (LSEC:DrawRowTitle rowIndex)

    (setq rowIndex
          (1+ rowIndex))

  )

  (LSEC:Print
    "Part 5F Row Titles Completed."
  )

  T

)


;-------------------------------------------------------------
; Part 5F Test
;-------------------------------------------------------------

(defun LSEC:TestRowTitles
       (/ rowIndex rowCount title pt)

  (LSEC:Print
    "----- Part 5F Row Title Geometry -----"
  )

  (LSEC:InitTableConfig)

  (setq rowIndex 0)

  (setq rowCount
        (LSEC:GetRowCount))

  (while (< rowIndex rowCount)

    (setq title
          (LSEC:GetRowTitle rowIndex))

    (setq pt
          (LSEC:GetRowTitlePoint rowIndex))

    (LSEC:Print
      (strcat
        "Row "
        (itoa rowIndex)
        " = "
        title
        " @ "
        (rtos (car pt) 2 2)
        ", "
        (rtos (cadr pt) 2 2)
      )
    )

    (setq rowIndex
          (1+ rowIndex))

  )

  (LSEC:Print
    "Part 5F Test Completed."
  )

  T

)

;=============================================================
;
; L-SECV6 Professional V6
; Part 5G - Station Columns / Chainage Labels
;
; Purpose:
;   Draw chainage labels in the DISTANCE IN MTRS. row.
;
; Important:
;   Part 5G uses the station geometry created in Part 5E.
;   It does NOT redraw the table frame.
;
;=============================================================


;-------------------------------------------------------------
; Part 5G.1 - Check Whether Value Exists In List
;
;-------------------------------------------------------------

(defun LSEC:ValueInList
       (value valueList / found item)

  (setq found nil)

  (foreach item valueList

    (if (= value item)
      (setq found T)
    )

  )

  found

)


;-------------------------------------------------------------
; Part 5G.2 - Get Display Station List
;
; Returns unique chainages while preserving their original order.
;
; Example:
;   0 20 40 288 288 300
;
; becomes:
;   0 20 40 288 300
;
;-------------------------------------------------------------

(defun LSEC:GetDisplayStationList
       (/ stations result station)

  (setq stations
        (LSEC:GetStationList))

  (setq result nil)

  (foreach station stations

    (if (not
          (LSEC:ValueInList
            station
            result
          )
        )

      (setq result
            (append
              result
              (list station)
            )
      )

    )

  )

  result

)


;-------------------------------------------------------------
; Part 5G.3 - Get Distance Row Index
;
; The DISTANCE IN MTRS. row is the final table row.
;
;-------------------------------------------------------------

(defun LSEC:GetDistanceRowIndex
       (/ rowCount)

  (setq rowCount
        (LSEC:GetRowCount))

  (if (> rowCount 0)

    (1- rowCount)

    0

  )

)


;-------------------------------------------------------------
; Part 5G.4 - Get Station Label Point
;
; X = station X from Part 5E
; Y = center of DISTANCE IN MTRS. row
;
;-------------------------------------------------------------

(defun LSEC:GetStationLabelPoint
       (chainage / rowIndex x y)

  (setq rowIndex
        (LSEC:GetDistanceRowIndex))

  (setq x
        (LSEC:GetStationCenter chainage))

  (setq y
        (LSEC:GetRowCenter rowIndex))

  (list x y 0.0)

)


;-------------------------------------------------------------
; Part 5G.5 - Format Station Label
;
; Whole-number chainages:
;   0
;   20
;   920
;
; Decimal chainages:
;   288.50
;
;-------------------------------------------------------------

(defun LSEC:FormatStationLabel
       (chainage)

  (if (= chainage (fix chainage))

    (itoa (fix chainage))

    (rtos chainage 2 2)

  )

)


;-------------------------------------------------------------
; Part 5G.6 - Draw One Station Label
;
; Draws vertical chainage text.
;
;-------------------------------------------------------------

(defun LSEC:DrawStationLabel
       (chainage / pt txt textHeight)

  (setq pt
        (LSEC:GetStationLabelPoint chainage))

  (setq txt
        (LSEC:FormatStationLabel chainage))

  (setq textHeight
        (LSEC:GetConfig "TEXTHEIGHT"))

  (if (null textHeight)

    (setq textHeight 3.5)

  )

  (entmakex
    (list
      '(0 . "TEXT")
      '(100 . "AcDbEntity")
      '(100 . "AcDbText")

      (cons 8 "LSEC-TABLE")

      (cons 10 pt)
      (cons 40 textHeight)
      (cons 1 txt)

      (cons 7 (getvar "TEXTSTYLE"))

      (cons 50 (/ pi 2.0))

      (cons 72 1)
      (cons 73 2)

      (cons 11 pt)
    )
  )

  T

)


;-------------------------------------------------------------
; Part 5G.7 - Draw All Station Labels
;
;-------------------------------------------------------------

(defun LSEC:DrawStationLabels
       (/ stations station)

  (LSEC:Print
    "Drawing Part 5G Station Labels..."
  )

  (LSEC:InitTableConfig)

  (LSEC:EnsureTableLayer)

  (setq stations
        (LSEC:GetDisplayStationList))

  (foreach station stations

    (LSEC:DrawStationLabel station)

  )

  (LSEC:Print
    "Part 5G Station Labels Completed."
  )

  T

)


;-------------------------------------------------------------
; Part 5G Test
;
;-------------------------------------------------------------

(defun LSEC:TestStationLabels
       (/ stations station pt)

  (LSEC:Print
    "----- Part 5G Station Label Geometry -----"
  )

  (LSEC:InitTableConfig)

  (setq stations
        (LSEC:GetDisplayStationList))

  (LSEC:Print
    (strcat
      "Display Station Count = "
      (itoa (length stations))
    )
  )

  (foreach station stations

    (setq pt
          (LSEC:GetStationLabelPoint station))

    (LSEC:Print
      (strcat
        "Station "
        (LSEC:FormatStationLabel station)
        " @ X = "
        (rtos (car pt) 2 2)
        " Y = "
        (rtos (cadr pt) 2 2)
      )
    )

  )

  (LSEC:Print
    "Part 5G Test Completed."
  )

  T

)

;=============================================================
;
; L-SECV6 Professional V6
; Part 5H - Populate Table Data
;
; Purpose:
;   Draw Excel/profile data inside the first six table rows.
;
; Row Mapping:
;
;   Row 0 = CUT DEPTH (m)        -> CUT
;   Row 1 = FILL DEPTH (m)       -> FILL
;   Row 2 = TOP BANK LEVEL (m)   -> BANK
;   Row 3 = FULL SUPPLY LEVEL (m)-> FSL
;   Row 4 = CANAL BED LEVEL (m)  -> BED
;   Row 5 = GROUND LEVEL (m)     -> EGL
;
; Row 6 is handled separately by Part 5G:
;   DISTANCE IN MTRS.
;
;=============================================================


;-------------------------------------------------------------
; Part 5H.1 - Get Data Field For Table Row
;-------------------------------------------------------------

(defun LSEC:GetTableDataKey
       (rowIndex)

  (cond

    ((= rowIndex 0) "CUT")
    ((= rowIndex 1) "FILL")
    ((= rowIndex 2) "BANK")
    ((= rowIndex 3) "FSL")
    ((= rowIndex 4) "BED")
    ((= rowIndex 5) "EGL")

    (T nil)

  )

)


;-------------------------------------------------------------
; Part 5H.2 - Get Table Data Value
;
; Returns the value from one loaded data record
; corresponding to the specified table row.
;
;-------------------------------------------------------------

(defun LSEC:GetTableDataValue
       (rowIndex rec / key item)

  (setq key
        (LSEC:GetTableDataKey rowIndex))

  (if key

    (progn

      (setq item
            (assoc key rec))

      (if item
        (cdr item)
        nil
      )

    )

    nil

  )

)


;-------------------------------------------------------------
; Part 5H.3 - Format Table Data Value
;
; Numeric values are displayed to 2 decimal places.
;
; Blank/nil values are not drawn.
;
;-------------------------------------------------------------

(defun LSEC:FormatTableDataValue
       (value)

  (cond

    ((null value)
      nil
    )

    ((numberp value)
      (rtos value 2 2)
    )

    (T
      (vl-princ-to-string value)
    )

  )

)


;-------------------------------------------------------------
; Part 5H.4 - Get Data Cell Point
;
; Uses Part 5E geometry:
;
; X = station center
; Y = row center
;
;-------------------------------------------------------------

(defun LSEC:GetDataCellPoint
       (rowIndex chainage)

  (LSEC:GetCellCenter
    rowIndex
    chainage
  )

)


;-------------------------------------------------------------
; Part 5H.5 - Draw One Table Data Value
;-------------------------------------------------------------

(defun LSEC:DrawTableDataValue
       (rowIndex rec / chainage value txt pt textHeight)

  (setq chainage
        (cdr
          (assoc "CHAINAGE" rec)))

  (setq value
        (LSEC:GetTableDataValue
          rowIndex
          rec))

  (setq txt
        (LSEC:FormatTableDataValue value))

  (if (and chainage txt)

    (progn

      (setq pt
            (LSEC:GetDataCellPoint
              rowIndex
              chainage))

      (setq textHeight
            (LSEC:GetConfig "TEXTHEIGHT"))

      (if (null textHeight)
        (setq textHeight 3.5)
      )

      (entmakex
        (list

          '(0 . "TEXT")
          '(100 . "AcDbEntity")
          '(100 . "AcDbText")

          (cons 8 "LSEC-TABLE")

          (cons 10 pt)
          (cons 40 textHeight)
          (cons 1 txt)

          (cons 7 (getvar "TEXTSTYLE"))

          ; Vertical text
          (cons 50 (/ pi 2.0))

          ; Center alignment
          (cons 72 1)
          (cons 73 2)

          (cons 11 pt)

        )
      )

    )

  )

  T

)

;-------------------------------------------------------------
; Part 5H.6 - Draw One Data Row
;
; Draws one complete table row for all stations.
;
;-------------------------------------------------------------

(defun LSEC:DrawTableDataRow
       (rowIndex / rec)

  (foreach rec *LSEC-DATA*

    (LSEC:DrawTableDataValue
      rowIndex
      rec
    )

  )

  T

)


;-------------------------------------------------------------
; Part 5H.7 - Draw All Table Data
;
; Draws rows 0 through 5.
;
; Row 6 belongs to Part 5G.
;
;-------------------------------------------------------------

(defun LSEC:DrawTableData
       (/ rowIndex)

  (LSEC:Print
    "Drawing Part 5H Table Data..."
  )

  (if (not *LSEC-DATA*)

    (progn

      (LSEC:Print
        "Part 5H aborted: no Excel data loaded."
      )

      nil

    )

    (progn

      (LSEC:InitTableConfig)

      (LSEC:EnsureTableLayer)

      (setq rowIndex 0)

      (while (< rowIndex 6)

        (LSEC:DrawTableDataRow
          rowIndex)

        (setq rowIndex
              (1+ rowIndex))

      )

      (LSEC:Print
        "Part 5H Table Data Completed."
      )

      T

    )

  )

)


;-------------------------------------------------------------
; Part 5H.8 - Test Table Data Mapping
;
; Tests the first and last records without drawing.
;
;-------------------------------------------------------------

(defun LSEC:TestTableData
       (/ firstRec lastRec rowIndex firstValue lastValue)

  (LSEC:Print
    "----- Part 5H Table Data Test -----"
  )

  (if (not *LSEC-DATA*)

    (progn

      (LSEC:Print
        "No Excel data loaded."
      )

      nil

    )

    (progn

      (setq firstRec
            (car *LSEC-DATA*))

      (setq lastRec
            (car (reverse *LSEC-DATA*)))

      (LSEC:Print
        (strcat
          "First Chainage = "
          (rtos
            (cdr (assoc "CHAINAGE" firstRec))
            2
            2
          )
        )
      )

      (LSEC:Print
        (strcat
          "Last Chainage = "
          (rtos
            (cdr (assoc "CHAINAGE" lastRec))
            2
            2
          )
        )
      )

      (setq rowIndex 0)

      (while (< rowIndex 6)

        (setq firstValue
              (LSEC:GetTableDataValue
                rowIndex
                firstRec))

        (setq lastValue
              (LSEC:GetTableDataValue
                rowIndex
                lastRec))

        (LSEC:Print
          (strcat
            "Row "
            (itoa rowIndex)
            " ["
            (LSEC:GetTableDataKey rowIndex)
            "] First = "
            (if
              (LSEC:FormatTableDataValue firstValue)
              (LSEC:FormatTableDataValue firstValue)
              "nil"
            )
            " Last = "
            (if
              (LSEC:FormatTableDataValue lastValue)
              (LSEC:FormatTableDataValue lastValue)
              "nil"
            )
          )
        )

        (setq rowIndex
              (1+ rowIndex))

      )

      (LSEC:Print
        "Part 5H Test Completed."
      )

      T

    )

  )

)

# Part 5I – Final Assembly


```lisp
;=============================================================
;
; L-SECV6 Professional V6
; Part 5I - Final Assembly
;
; Purpose:
;   Assemble all completed drawing components into one
;   longitudinal section drawing.
;
; Drawing Order:
;
;   1. Calculate DATUM
;   2. Sheet Frame
;   3. Profiles
;   4. Grid
;   5. Structures
;   6. Main Title
;   7. Title Block
;   8. Table Frame
;   9. Row Titles
;  10. Station Labels
;  11. Table Data
;
;=============================================================


;-------------------------------------------------------------
; Part 5I.1 - Draw Single Longitudinal Section
;-------------------------------------------------------------

(defun LSEC:DrawSingleSection
       (/ result)

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "L-SECV6 Professional - Drawing Started"
  )

  (LSEC:Print
    "========================================"
  )

  ;-----------------------------------------------------------
  ; Validate Excel Data
  ;-----------------------------------------------------------

  (if (not *LSEC-DATA*)

    (progn

      (LSEC:Print
        "Drawing aborted: no Excel data loaded."
      )

      nil

    )

    (progn

      (LSEC:InitConfig)
      (LSEC:InitLayers)
      (LSEC:InitTableConfig)

      (LSEC:Print
        "1/11 Calculating datum..."
      )

      (LSEC:CalculateDatum)

      (LSEC:Print
        "2/11 Drawing sheet frame..."
      )

      (LSEC:DrawSheetFrame)

      (LSEC:Print
        "3/11 Drawing profiles..."
      )

      (LSEC:DrawProfiles)

      (LSEC:Print
        "4/11 Drawing grid..."
      )

      (LSEC:DrawGrid)

      (LSEC:Print
        "5/11 Drawing structures..."
      )

      (LSEC:DrawStructures)

      (LSEC:Print
        "Drawing Part 5J profile legend..."
      )

      (LSEC:DrawLegend)

      (LSEC:Print
        "6/11 Drawing title..."
      )

      (LSEC:DrawTitle)

      (LSEC:Print
        "7/11 Drawing title block..."
      )

      (LSEC:DrawTitleBlock)

      (LSEC:Print
        "8/11 Drawing table frame..."
      )

      (LSEC:DrawTableFrame)

      (LSEC:Print
        "9/11 Drawing row titles..."
      )

      (LSEC:DrawRowTitles)

      (LSEC:Print
        "10/11 Drawing station labels..."
      )

      (LSEC:DrawStationLabels)

      (LSEC:Print
        "11/11 Drawing table data..."
      )

      (LSEC:DrawTableData)

      (LSEC:Print
        "========================================"
      )

      (LSEC:Print
        "L-SECV6 Drawing Completed Successfully"
      )

      (LSEC:Print
        "========================================"
      )

      T

    )

  )

)

;=============================================================
;
; Part 5J.9 - Final Presentation Calibration Override
; Split-Aware Version
;
;=============================================================


;-------------------------------------------------------------
; Lower the drawing datum to give the profile more clear space.
;-------------------------------------------------------------

(defun LSEC:CalculateDatum (/ minElev offset)

  (setq minElev
    (LSEC:FindMinElevation)
  )

  (if minElev

    (progn

      (setq offset 3.0)

      (setq *LSEC-DATUM*
        (- minElev offset)
      )

      (LSEC:SetConfig
        "DATUM"
        *LSEC-DATUM*
      )

      (LSEC:Print
        (strcat
          "Datum calculated: "
          (rtos *LSEC-DATUM* 2 3)
        )
      )

      *LSEC-DATUM*

    )

    nil

  )

)


;-------------------------------------------------------------
; Remove the large outer sheet/frame box.
;-------------------------------------------------------------

(defun LSEC:DrawSheetFrame ()

  (LSEC:Print
    "Sheet frame suppressed for clean reference-style presentation."
  )

  T

)


;-------------------------------------------------------------
; Remove profile grid lines.
;
; Keep only:
;   - Beginning elevation axis
;   - Elevation ticks
;   - RL labels
;   - Beginning/end profile boundaries
;-------------------------------------------------------------

(defun LSEC:DrawGrid (/)

  (LSEC:Print
    "Drawing clean profile boundaries..."
  )

  (LSEC:DrawProfileBoundaries)

  (LSEC:Print
    "Profile grid suppressed."
  )

  T

)


;=============================================================
;
; Split-Aware Profile Boundaries and Elevation Axis
;
; IMPORTANT:
; Use ActiveCHtoX so that the elevation axis,
; elevation labels and boundary lines move with
; each split section.
;
;=============================================================

(defun LSEC:DrawProfileBoundaries
       (/ firstRec
          lastRec
          start
          end
          x1
          x2
          datum
          maxElev
          y1
          y2
          tickLen
          elev
          step
          y)

  (if *LSEC-DATA*

    (progn

      (setq firstRec
        (car *LSEC-DATA*)
      )

      (setq lastRec
        (car (reverse *LSEC-DATA*))
      )

      (setq start
        (cdr (assoc "CHAINAGE" firstRec))
      )

      (setq end
        (cdr (assoc "CHAINAGE" lastRec))
      )

      (setq datum
        (LSEC:GetConfig "DATUM")
      )

      (setq maxElev
        (LSEC:FindMaxElevation)
      )

      (setq step
        (LSEC:GetGridY)
      )

      (setq tickLen 4.0)

      (if
        (and
          (numberp start)
          (numberp end)
          (numberp datum)
          (numberp maxElev)
          (numberp step)
          (> step 0.0)
        )

        (progn

          ;---------------------------------------------------
          ; Split-aware X coordinates.
          ;---------------------------------------------------

          (setq x1
            (LSEC:ActiveCHtoX start)
          )

          (setq x2
            (LSEC:ActiveCHtoX end)
          )

          (setq y1
            (LSEC:ELtoY datum)
          )

          (setq y2
            (LSEC:ELtoY (+ maxElev 5.0))
          )


          ; Beginning elevation axis.
          (LSEC:DrawLine
            (list x1 y1 0.0)
            (list x1 y2 0.0)
            "LSEC-GRID"
          )


          ; End boundary.
          (LSEC:DrawLine
            (list x2 y1 0.0)
            (list x2 y2 0.0)
            "LSEC-GRID"
          )


          ; Elevation ticks and RL labels.
          (setq elev datum)

          (while (<= elev maxElev)

            (setq y
              (LSEC:ELtoY elev)
            )

            (LSEC:DrawLine
              (list (- x1 tickLen) y 0.0)
              (list x1 y 0.0)
              "LSEC-GRID"
            )

            (LSEC:DrawText
              (list
                (- x1 18.0)
                (- y 1.0)
                0.0
              )
              (rtos elev 2 2)
              (LSEC:GetTextHeight)
              "LSEC-TEXT"
            )

            (setq elev
              (+ elev step)
            )

          )


          ; Elevation title.
          (LSEC:DrawTextRotated
            (list
              (- x1 30.0)
              (/ (+ y1 y2) 2.0)
              0.0
            )
            "ELEVATION (m)"
            (LSEC:GetTextHeight)
            90.0
            "LSEC-TEXT"
          )


          (LSEC:Print
            "Split-aware profile boundaries and elevation axis completed."
          )

          T

        )

      )

    )

  )

)


;=============================================================
;
; Split-Aware Profile Legend
;
;=============================================================

(defun LSEC:DrawLegend
       (/ firstRec
          datum
          x1
          x
          y
          lineLen
          gap
          h
          items
          item)

  (if *LSEC-DATA*

    (progn

      (setq firstRec
        (car *LSEC-DATA*)
      )


      ;-------------------------------------------------------
      ; IMPORTANT:
      ; Split-aware section X position.
      ;-------------------------------------------------------

      (setq x1
        (LSEC:ActiveCHtoX
          (cdr (assoc "CHAINAGE" firstRec))
        )
      )

      (setq datum
        (LSEC:GetConfig "DATUM")
      )


      ; Put legend clearly outside and left of the section.
      (setq x
        (- x1 95.0)
      )

      (setq y
        (+ (LSEC:ELtoY datum) 28.0)
      )

      (setq lineLen 28.0)
      (setq gap 6.0)
      (setq h (LSEC:GetTextHeight))


      (setq items

        (list

          (list "LSEC-EGL"  "GROUND LEVEL")
          (list "LSEC-BANK" "TOP BANK LEVEL")
          (list "LSEC-FSL"  "FULL SUPPLY LEVEL")
          (list "LSEC-BED"  "CANAL BED LEVEL")

        )

      )


      (foreach item items

        (LSEC:DrawLine
          (list x y 0.0)
          (list (+ x lineLen) y 0.0)
          (car item)
        )

        (LSEC:DrawText
          (list
            (+ x lineLen 5.0)
            (- y 1.25)
            0.0
          )
          (cadr item)
          h
          "LSEC-TEXT"
        )

        (setq y
          (- y gap)
        )

      )


      (LSEC:Print
        "Split-aware profile legend completed."
      )

      T

    )

  )

)


;=============================================================
;
; Split-Aware Structures
;
;=============================================================

(defun LSEC:DrawStructures
       (/ rec
          ch
          structName
          x
          egl
          bank
          fsl
          bed
          topElev
          yBase
          flagHeight
          textOffset
          elev)

  (LSEC:CreateLayer
    "LSEC-STRUCT"
    1
  )


  ; Higher presentation clearance.
  (setq flagHeight 35.0)
  (setq textOffset 6.0)


  (foreach rec *LSEC-DATA*

    (setq structName
      (cdr (assoc "STRUCT" rec))
    )


    (if
      (and
        structName
        (/= structName "")
      )

      (progn

        (setq ch
          (cdr (assoc "CHAINAGE" rec))
        )

        (setq egl
          (cdr (assoc "EGL" rec))
        )

        (setq bank
          (cdr (assoc "BANK" rec))
        )

        (setq fsl
          (cdr (assoc "FSL" rec))
        )

        (setq bed
          (cdr (assoc "BED" rec))
        )


        ;-----------------------------------------------------
        ; Find highest visible profile elevation.
        ;-----------------------------------------------------

        (setq topElev nil)

        (foreach elev
          (list egl bank fsl bed)

          (if
            (and
              (numberp elev)
              (or
                (not topElev)
                (> elev topElev)
              )
            )

            (setq topElev elev)

          )

        )


        ;-----------------------------------------------------
        ; IMPORTANT:
        ; Split-aware X coordinate.
        ;-----------------------------------------------------

        (setq x
          (LSEC:ActiveCHtoX ch)
        )


        (setq yBase

          (if topElev
            (LSEC:ELtoY topElev)
            nil
          )

        )


        (if
          (and
            (numberp x)
            (numberp yBase)
          )

          (progn


            ; Flag starts directly on the profile line.
            (LSEC:DrawLine
              (list x yBase 0.0)
              (list
                x
                (+ yBase flagHeight)
                0.0
              )
              "LSEC-STRUCT"
            )


            ; Structure label above flag.
            (LSEC:DrawTextRotated
              (list
                (+ x 4.0)
                (+ yBase flagHeight textOffset)
                0.0
              )
              structName
              (LSEC:GetTextHeight)
              65.0
              "LSEC-TEXT"
            )

          )

        )

      )

    )

  )


  (LSEC:Print
    "Split-aware structures completed with raised white labels."
  )

  T

)


(LSEC:Print
  "Part 5J.9 Final Presentation Calibration Override Loaded."
)

;=============================================================
;
; Part 6A - Split Range Engine
;
;=============================================================

;-------------------------------------------------------------
; Get Minimum Chainage
;-------------------------------------------------------------

(defun LSEC:GetMinChainage (/ rec ch minCH)

  (setq minCH nil)

  (foreach rec *LSEC-DATA*

    (setq ch
      (cdr (assoc "CHAINAGE" rec))
    )

    (if (numberp ch)

      (if (or (not minCH)
              (< ch minCH)
          )

        (setq minCH ch)

      )

    )

  )

  minCH

)


;-------------------------------------------------------------
; Get Maximum Chainage
;-------------------------------------------------------------

(defun LSEC:GetMaxChainage (/ rec ch maxCH)

  (setq maxCH nil)

  (foreach rec *LSEC-DATA*

    (setq ch
      (cdr (assoc "CHAINAGE" rec))
    )

    (if (numberp ch)

      (if (or (not maxCH)
              (> ch maxCH)
          )

        (setq maxCH ch)

      )

    )

  )

  maxCH

)


;-------------------------------------------------------------
; Get Overall Chainage Range
;
; Returns:
; (startCH endCH)
;
; Example:
; (0.0 2350.0)
;-------------------------------------------------------------

(defun LSEC:GetOverallChainageRange (/ minCH maxCH)

  (setq minCH (LSEC:GetMinChainage))
  (setq maxCH (LSEC:GetMaxChainage))

  (if (and (numberp minCH)
           (numberp maxCH)
      )

    (list minCH maxCH)

    nil

  )

)


;-------------------------------------------------------------
; Calculate Split Ranges
;
; Example:
;
; Start = 0
; End   = 2350
; Split = 500
;
; Result:
;
; (
;   (0 500)
;   (500 1000)
;   (1000 1500)
;   (1500 2000)
;   (2000 2350)
; )
;
;-------------------------------------------------------------

(defun LSEC:CalculateSplitRanges
       (/ range startCH endCH splitLength
          currentCH nextCH ranges)

  (setq range
    (LSEC:GetOverallChainageRange)
  )

  (setq splitLength
    (LSEC:GetConfig "SPLITLENGTH")
  )

  (if (and range
           (numberp splitLength)
           (> splitLength 0.0)
      )

    (progn

      (setq startCH
        (car range)
      )

      (setq endCH
        (cadr range)
      )

      (setq currentCH startCH)
      (setq ranges nil)

      (while (< currentCH endCH)

        (setq nextCH
          (+ currentCH splitLength)
        )

        (if (> nextCH endCH)

          (setq nextCH endCH)

        )

        (setq ranges
          (append
            ranges
            (list
              (list currentCH nextCH)
            )
          )
        )

        (setq currentCH nextCH)

      )

      ranges

    )

    nil

  )

)


;-------------------------------------------------------------
; Get Active Split Ranges
;
; SPLITMODE = 0
; Returns one complete range:
; ((startCH endCH))
;
; SPLITMODE = 1
; Returns calculated split ranges.
;-------------------------------------------------------------

(defun LSEC:GetSplitRanges
       (/ range)

  (if (= (LSEC:GetConfig "SPLITMODE") 1)

    (LSEC:CalculateSplitRanges)

    (progn

      (setq range
        (LSEC:GetOverallChainageRange)
      )

      (if range
        (list range)
        nil
      )

    )

  )

)


;-------------------------------------------------------------
; Test Split Range Engine
;-------------------------------------------------------------

(defun LSEC:TestSplitRanges
       (/ range ranges index)

  (if (not *LSEC-DATA*)

    (LSEC:Print
      "No Excel data loaded."
    )

    (progn

      (setq range
        (LSEC:GetOverallChainageRange)
      )

      (LSEC:Print
        (strcat
          "Overall chainage: "
          (rtos (car range) 2 2)
          " to "
          (rtos (cadr range) 2 2)
        )
      )

      (LSEC:Print
        (strcat
          "Split mode: "
          (itoa
            (LSEC:GetConfig "SPLITMODE")
          )
        )
      )

      (LSEC:Print
        (strcat
          "Split length: "
          (rtos
            (LSEC:GetConfig "SPLITLENGTH")
            2
            2
          )
        )
      )

      (setq ranges
        (LSEC:GetSplitRanges)
      )

      (setq index 1)

      (foreach range ranges

        (LSEC:Print
          (strcat
            "Section "
            (itoa index)
            ": "
            (rtos (car range) 2 2)
            " to "
            (rtos (cadr range) 2 2)
          )
        )

        (setq index
          (1+ index)
        )

      )

      ranges

    )

  )

)

;=============================================================
;
; Part 6B - Section Data Filtering
;
;=============================================================

;-------------------------------------------------------------
; Get Data Records for a Section Range
;
; startCH = section starting chainage
; endCH   = section ending chainage
;
; Both boundary stations are included.
;
; Example:
; (LSEC:GetSectionData 0.0 200.0)
;
;-------------------------------------------------------------

(defun LSEC:GetSectionData
       (startCH endCH / rec ch sectionData)

  (setq sectionData nil)

  (if (and *LSEC-DATA*
           (numberp startCH)
           (numberp endCH)
           (<= startCH endCH))

    (progn

      (foreach rec *LSEC-DATA*

        (setq ch
          (cdr (assoc "CHAINAGE" rec))
        )

        ;; Include both section boundaries.
        (if (and (numberp ch)
                 (>= ch startCH)
                 (<= ch endCH))

          (setq sectionData
            (append sectionData (list rec))
          )

        )

      )

      sectionData

    )

    nil

  )

)


;-------------------------------------------------------------
; Set Active Section Data
;
; Stores the selected records in:
;
; *LSEC-SECTION-DATA*
;
;-------------------------------------------------------------

(defun LSEC:SetSectionData
       (startCH endCH / sectionData)

  (setq sectionData
    (LSEC:GetSectionData startCH endCH)
  )

  (setq *LSEC-SECTION-DATA*
    sectionData
  )

  (LSEC:Print
    (strcat
      "Active section data set: "
      (rtos startCH 2 2)
      " to "
      (rtos endCH 2 2)
      " | Records: "
      (itoa (length *LSEC-SECTION-DATA*))
    )
  )

  *LSEC-SECTION-DATA*

)


;-------------------------------------------------------------
; Get Active Drawing Data
;
; During normal drawing:
;     returns *LSEC-DATA*
;
; During split drawing:
;     returns *LSEC-SECTION-DATA*
;
; This will later allow existing Parts 4–5 to use
; the correct data without redesigning every function.
;
;-------------------------------------------------------------

(defun LSEC:GetActiveData (/)

  (if (and (boundp '*LSEC-SECTION-DATA*)
           *LSEC-SECTION-DATA*)

    *LSEC-SECTION-DATA*

    *LSEC-DATA*

  )

)


;-------------------------------------------------------------
; Clear Active Section Data
;
; Returns drawing control to the complete Excel data.
;
;-------------------------------------------------------------

(defun LSEC:ClearSectionData (/)

  (setq *LSEC-SECTION-DATA* nil)

  (LSEC:Print
    "Active section data cleared."
  )

  T

)


;-------------------------------------------------------------
; Test Section Data
;
; Example:
;
; (LSEC:TestSectionData 0.0 200.0)
;
;-------------------------------------------------------------

(defun LSEC:TestSectionData
       (startCH endCH / sectionData firstRec lastRec)

  (setq sectionData
    (LSEC:GetSectionData startCH endCH)
  )

  (if sectionData

    (progn

      (setq firstRec
        (car sectionData)
      )

      (setq lastRec
        (car (reverse sectionData))
      )

      (LSEC:Print
        (strcat
          "Section range: "
          (rtos startCH 2 2)
          " to "
          (rtos endCH 2 2)
        )
      )

      (LSEC:Print
        (strcat
          "Records selected: "
          (itoa (length sectionData))
        )
      )

      (LSEC:Print
        (strcat
          "First CH: "
          (rtos
            (cdr (assoc "CHAINAGE" firstRec))
            2
            2
          )
        )
      )

      (LSEC:Print
        (strcat
          "Last CH: "
          (rtos
            (cdr (assoc "CHAINAGE" lastRec))
            2
            2
          )
        )
      )

      sectionData

    )

    (progn

      (LSEC:Print
        "No records found for section."
      )

      nil

    )

  )

)


(LSEC:Print "Part 6B - Section Data Filtering Loaded.")

;=============================================================
;
; Part 6C - Active Section Context
;
;=============================================================


;-------------------------------------------------------------
; Section Context Variables
;
; *LSEC-SAVED-DATA*  = Original complete Excel data
; *LSEC-SECTION-START* = Active section start chainage
; *LSEC-SECTION-END*   = Active section end chainage
;
;-------------------------------------------------------------

(setq *LSEC-SAVED-DATA* nil)
(setq *LSEC-SECTION-START* nil)
(setq *LSEC-SECTION-END* nil)


;-------------------------------------------------------------
; Activate Section
;
; Saves complete data and temporarily makes the selected
; section the active drawing data.
;
; Example:
;
; (LSEC:ActivateSection 0.0 200.0)
;
;-------------------------------------------------------------

(defun LSEC:ActivateSection
       (startCH endCH / sectionData)

  (setq sectionData
    (LSEC:GetSectionData startCH endCH)
  )

  (if sectionData

    (progn

      ;; Save the complete Excel data only once.
      (if (not *LSEC-SAVED-DATA*)
        (setq *LSEC-SAVED-DATA*
          *LSEC-DATA*
        )
      )

      ;; Store section information.
      (setq *LSEC-SECTION-DATA*
        sectionData
      )

      (setq *LSEC-SECTION-START*
        startCH
      )

      (setq *LSEC-SECTION-END*
        endCH
      )

      ;; Make the section temporarily become the active
      ;; data source for the existing drawing functions.
      (setq *LSEC-DATA*
        sectionData
      )

      (LSEC:Print
        (strcat
          "Section activated: "
          (rtos startCH 2 2)
          " to "
          (rtos endCH 2 2)
          " | Records: "
          (itoa (length sectionData))
        )
      )

      T

    )

    (progn

      (LSEC:Print
        "Section activation failed: no records found."
      )

      nil

    )

  )

)


;-------------------------------------------------------------
; Deactivate Section
;
; Restores the complete Excel data.
;
;-------------------------------------------------------------

(defun LSEC:DeactivateSection (/)

  (if *LSEC-SAVED-DATA*

    (progn

      (setq *LSEC-DATA*
        *LSEC-SAVED-DATA*
      )

      (setq *LSEC-SAVED-DATA*
        nil
      )

      (setq *LSEC-SECTION-DATA*
        nil
      )

      (setq *LSEC-SECTION-START*
        nil
      )

      (setq *LSEC-SECTION-END*
        nil
      )

      (LSEC:Print
        "Complete Excel data restored."
      )

      T

    )

    (progn

      (LSEC:Print
        "No active section to deactivate."
      )

      nil

    )

  )

)


;-------------------------------------------------------------
; Get Active Section Range
;
; Returns:
;
; (startCH endCH)
;
; or nil when no section is active.
;
;-------------------------------------------------------------

(defun LSEC:GetActiveSectionRange (/)

  (if (and (numberp *LSEC-SECTION-START*)
           (numberp *LSEC-SECTION-END*))

    (list
      *LSEC-SECTION-START*
      *LSEC-SECTION-END*
    )

    nil

  )

)


;-------------------------------------------------------------
; Is Section Active?
;
; Returns T or nil.
;
;-------------------------------------------------------------

(defun LSEC:SectionActiveP (/)

  (if (and *LSEC-SAVED-DATA*
           *LSEC-SECTION-DATA*
           (numberp *LSEC-SECTION-START*)
           (numberp *LSEC-SECTION-END*))

    T

    nil

  )

)


;-------------------------------------------------------------
; Test Active Section Context
;
; Example:
;
; (LSEC:TestActivateSection 0.0 200.0)
;
;-------------------------------------------------------------

(defun LSEC:TestActivateSection
       (startCH endCH / firstRec lastRec)

  (if (LSEC:ActivateSection startCH endCH)

    (progn

      (setq firstRec
        (car *LSEC-DATA*)
      )

      (setq lastRec
        (car (reverse *LSEC-DATA*))
      )

      (LSEC:Print
        (strcat
          "Active first CH: "
          (rtos
            (cdr (assoc "CHAINAGE" firstRec))
            2
            2
          )
        )
      )

      (LSEC:Print
        (strcat
          "Active last CH: "
          (rtos
            (cdr (assoc "CHAINAGE" lastRec))
            2
            2
          )
        )
      )

      (LSEC:Print
        (strcat
          "Active records: "
          (itoa (length *LSEC-DATA*))
        )
      )

      T

    )

    nil

  )

)


;-------------------------------------------------------------
; Test Restore Complete Data
;
;-------------------------------------------------------------

(defun LSEC:TestDeactivateSection (/)

  (LSEC:DeactivateSection)

  (if *LSEC-DATA*

    (LSEC:Print
      (strcat
        "Complete data records: "
        (itoa (length *LSEC-DATA*))
      )
    )

  )

  T

)


(LSEC:Print "Part 6C - Active Section Context Loaded.")

;=============================================================
;
; Part 6D - Draw Complete Section From Range
;
;=============================================================


;-------------------------------------------------------------
; Draw One Section
;
; Example:
;
; (LSEC:DrawSection 0.0 200.0)
;
;-------------------------------------------------------------

(defun LSEC:DrawSection
       (startCH endCH / result)

  (LSEC:Print "========================================")
  (LSEC:Print
    (strcat
      "Drawing section: "
      (rtos startCH 2 2)
      " to "
      (rtos endCH 2 2)
    )
  )
  (LSEC:Print "========================================")

  ;; Activate selected section.
  (if (LSEC:ActivateSection startCH endCH)

    (progn

      ;; Draw using the existing drawing engine.
      (setq result
        (LSEC:DrawCompleteSection)
      )

      ;; Always restore the complete Excel data.
      (LSEC:DeactivateSection)

      (if result

        (LSEC:Print
          (strcat
            "Section completed: "
            (rtos startCH 2 2)
            " to "
            (rtos endCH 2 2)
          )
        )

        (LSEC:Print
          "Section drawing failed."
        )

      )

      result

    )

    (progn

      (LSEC:Print
        "Section drawing aborted: unable to activate section."
      )

      nil

    )

  )

)


;-------------------------------------------------------------
; Test Draw First Split Section
;
; Automatically gets the first range from the split engine.
;
;-------------------------------------------------------------

(defun LSEC:TestDrawFirstSection
       (/ ranges firstRange)

  (setq ranges
    (LSEC:GetSplitRanges)
  )

  (if ranges

    (progn

      (setq firstRange
        (car ranges)
      )

      (LSEC:DrawSection
        (car firstRange)
        (cadr firstRange)
      )

    )

    (progn

      (LSEC:Print
        "No split ranges available."
      )

      nil

    )

  )

)


(LSEC:Print
  "Part 6D - Draw Complete Section From Range Loaded."
)

;=============================================================
;
; Part 6E - Section Drawing Width
;
;=============================================================


;-------------------------------------------------------------
; Get Section Drawing Width
;
; Converts the actual chainage length of one section into
; drawing X-units.
;
; Example:
;
; Section:
;   200.00 to 400.00
;
; Local chainage length:
;   400 - 200 = 200
;
; Drawing width:
;   CHtoX(200)
;
;-------------------------------------------------------------

(defun LSEC:GetSectionDrawingWidth
       (startCH endCH / sectionLength width)

  (if
    (and
      (numberp startCH)
      (numberp endCH)
      (>= endCH startCH)
    )

    (progn

      ;-------------------------------------------------------
      ; Calculate actual section chainage length.
      ;-------------------------------------------------------

      (setq sectionLength
        (- endCH startCH)
      )


      ;-------------------------------------------------------
      ; Convert local section length to drawing units.
      ;
      ; IMPORTANT:
      ; Do NOT use ActiveCHtoX here.
      ;
      ; ActiveCHtoX includes the frame offset.
      ; This function must return ONLY the width of
      ; the section itself.
      ;-------------------------------------------------------

      (setq width
        (LSEC:CHtoX sectionLength)
      )


      (if (numberp width)
        width
        0.0
      )

    )

    0.0

  )

)

;-------------------------------------------------------------
; Part 6E - Calculate Section Frames
;
; frameX = X coordinate of the first PROFILE station.
;
; Each table extends TABLELABELWIDTH to the left of frameX.
; Therefore labelWidth must be reserved before the next frame.
;
;-------------------------------------------------------------

(defun LSEC:CalculateSectionFrames
       (/ ranges
          gap
          labelWidth
          frameX
          range
          startCH
          endCH
          width
          result)

  (setq ranges
    (LSEC:GetSplitRanges)
  )

  ; Get required visible gap between table frames.
  (setq gap
    (LSEC:GetConfig "SECTIONGAP")
  )

  (if (not (numberp gap))
    (setq gap 50.0)
  )

  ; Make sure table configuration exists.
  (LSEC:InitTableConfig)

  ; Each new table has its own label column.
  (setq labelWidth
    (LSEC:GetLabelColumnWidth)
  )

  (if (not (numberp labelWidth))
    (setq labelWidth 0.0)
  )

  ; First section profile begins at X = 0.
  (setq frameX 0.0)

  (setq result nil)

  (foreach range ranges

    (setq startCH (car range))
    (setq endCH   (cadr range))

    ; Width of the actual profile portion.
    (setq width
      (LSEC:GetSectionDrawingWidth
        startCH
        endCH
      )
    )

    ; Store:
    ; (startCH endCH frameX)
    (setq result
      (append
        result
        (list
          (list startCH endCH frameX)
        )
      )
    )

    ;---------------------------------------------------------
    ; Move to the next section.
    ;
    ; width      = current profile width
    ; labelWidth = repeated label column of next table
    ; gap        = visible table-to-table gap
    ;---------------------------------------------------------

    (setq frameX
      (+
        frameX
        width
        labelWidth
        gap
      )
    )

  )

  result

)


;-------------------------------------------------------------
; Test Section Frame Placement
;
;-------------------------------------------------------------

(defun LSEC:TestSectionFrames
       (/ frames frame index startCH endCH frameX width)

  (setq frames
    (LSEC:CalculateSectionFrames)
  )

  (setq index 1)

  (foreach frame frames

    (setq startCH (car frame))
    (setq endCH (cadr frame))
    (setq frameX (caddr frame))

    (setq width
      (LSEC:GetSectionDrawingWidth
        startCH endCH
      )
    )

    (LSEC:Print
      (strcat
        "Section "
        (itoa index)
        ": CH "
        (rtos startCH 2 2)
        " - "
        (rtos endCH 2 2)
        " | Frame X = "
        (rtos frameX 2 2)
        " | Width = "
        (rtos width 2 2)
      )
    )

    (setq index (1+ index))

  )

  frames

)


(LSEC:Print
  "Part 6E - Split Section Frame Placement Loaded."
)

;=============================================================
;
; Part 6F - Split Section Drawing Integration
; Safe Version
;
;=============================================================

;-------------------------------------------------------------
; Active Chainage to Drawing X
;
; Normal drawing:
;   original CHtoX(CH)
;
; Split drawing:
;   Frame X + original CHtoX(CH - Section Start)
;
;-------------------------------------------------------------

(defun LSEC:ActiveCHtoX
       (ch / localCH)

  (if (numberp ch)

    (if (and (numberp *LSEC-SECTION-CHAINAGE-START*)
             (numberp *LSEC-SECTION-FRAME-X*))

      (progn

        (setq localCH
          (- ch *LSEC-SECTION-CHAINAGE-START*)
        )

        (+ *LSEC-SECTION-FRAME-X*
           (LSEC:CHtoX localCH)
        )

      )

      (LSEC:CHtoX ch)

    )

    nil

  )

)


;-------------------------------------------------------------
; Activate Split Frame
;-------------------------------------------------------------

(defun LSEC:ActivateSectionFrame
       (startCH frameX)

  (setq *LSEC-SECTION-CHAINAGE-START* startCH)
  (setq *LSEC-SECTION-FRAME-X* frameX)

  T

)


;-------------------------------------------------------------
; Deactivate Split Frame
;-------------------------------------------------------------

(defun LSEC:DeactivateSectionFrame ()

  (setq *LSEC-SECTION-CHAINAGE-START* nil)
  (setq *LSEC-SECTION-FRAME-X* nil)

  T

)


;-------------------------------------------------------------
; Draw One Section At Frame
;-------------------------------------------------------------

(defun LSEC:DrawSectionAtFrame
       (startCH endCH frameX / result)

  (if (LSEC:ActivateSection startCH endCH)

    (progn

      (LSEC:ActivateSectionFrame
        startCH
        frameX
      )

      (setq result
        (LSEC:DrawSingleSection)
      )

      (LSEC:DeactivateSectionFrame)

      (LSEC:DeactivateSection)

      result

    )

    nil

  )

)


;-------------------------------------------------------------
; Draw All Split Sections
;-------------------------------------------------------------

(defun LSEC:DrawAllSplitSections
       (/ frames frame startCH endCH frameX result)

  (setq frames
    (LSEC:CalculateSectionFrames)
  )

  (setq result T)

  (foreach frame frames

    (setq startCH (car frame))
    (setq endCH   (cadr frame))
    (setq frameX  (caddr frame))

    (if
      (not
        (LSEC:DrawSectionAtFrame
          startCH
          endCH
          frameX
        )
      )

      (setq result nil)

    )

  )

  result

)


(LSEC:Print
  "Part 6F - Split Section Drawing Integration Loaded."
)

;=============================================================
;
; Part 6G - Drawing Dispatcher
;
;=============================================================

;-------------------------------------------------------------
; Draw Complete Section
;
; SPLITMODE = 0
;   Draw one complete longitudinal section.
;
; SPLITMODE = 1
;   Draw all calculated split sections.
;
;-------------------------------------------------------------

(defun LSEC:DrawCompleteSection
       (/ splitMode result)

  ; Ensure configuration exists.
  (LSEC:InitConfig)

  (setq splitMode
        (LSEC:GetConfig "SPLITMODE"))

  (if (= splitMode 1)

    (progn

      (LSEC:Print
        "Split drawing mode activated."
      )

      (setq result
            (LSEC:DrawAllSplitSections))

      result

    )

    (progn

      (LSEC:Print
        "Normal single-section drawing mode activated."
      )

      (setq result
            (LSEC:DrawSingleSection))

      result

    )

  )

)


(LSEC:Print
  "Part 6G - Drawing Dispatcher Loaded."
)

;=============================================================
;
; Part 7A - Final User Drawing Command
; Updated Version with Part 7B and Part 7C Integration
;
; Command:
;   LSECFINALDRAW
;
; Workflow:
;   1. Prompt for Excel file
;   2. Load Excel data automatically
;   3. Prompt Horizontal Scale
;   4. Prompt Vertical Exaggeration
;   5. Prompt Split Length
;   6. Prepare final drawing
;   7. Draw all split sections
;
;=============================================================


(defun c:LSECFINALDRAW
       (/ hscale
          vexag
          splitLength
          defaultHScale
          defaultVExag
          defaultSplitLength
          result)

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "L-SECV6 Professional - Final Drawing"
  )

  (LSEC:Print
    "========================================"
  )


  ;-----------------------------------------------------------
  ; Part 7C - Select and Load Excel File
  ;-----------------------------------------------------------

  (if (LSEC:PromptAndLoadExcel)

    (progn


      ;-------------------------------------------------------
      ; Get current/default Horizontal Scale
      ;-------------------------------------------------------

      (setq defaultHScale
        (LSEC:GetConfig "HSCALE")
      )

      (if (not (numberp defaultHScale))
        (setq defaultHScale 1000.0)
      )


      ;-------------------------------------------------------
      ; Get current/default Vertical Exaggeration
      ;-------------------------------------------------------

      (setq defaultVExag
        (LSEC:GetConfig "VEXAG")
      )

      (if (not (numberp defaultVExag))
        (setq defaultVExag 10.0)
      )


      ;-------------------------------------------------------
      ; Get current/default Split Length
      ;-------------------------------------------------------

      (setq defaultSplitLength
        (LSEC:GetConfig "SPLITLENGTH")
      )

      (if (not (numberp defaultSplitLength))
        (setq defaultSplitLength 200.0)
      )


      ;-------------------------------------------------------
      ; Prompt Horizontal Scale
      ;-------------------------------------------------------

      (setq hscale
        (LSEC:GetPositiveReal
          "Enter Horizontal Scale"
          defaultHScale
        )
      )


      ;-------------------------------------------------------
      ; Prompt Vertical Exaggeration
      ;-------------------------------------------------------

      (setq vexag
        (LSEC:GetPositiveReal
          "Enter Vertical Exaggeration"
          defaultVExag
        )
      )


      ;-------------------------------------------------------
      ; Prompt Split Length
      ;-------------------------------------------------------

      (setq splitLength
        (LSEC:GetPositiveReal
          "Enter Split Length"
          defaultSplitLength
        )
      )


      ;-------------------------------------------------------
      ; Apply User Settings
      ;-------------------------------------------------------

      (LSEC:SetConfig
        "HSCALE"
        hscale
      )

      (LSEC:SetConfig
        "VEXAG"
        vexag
      )

      (LSEC:SetConfig
        "SPLITLENGTH"
        splitLength
      )


      ;-------------------------------------------------------
      ; Apply Automatic Presentation Settings
      ;
      ; These are not user-prompted.
      ;-------------------------------------------------------

      (LSEC:SetConfig
        "SPLITMODE"
        1
      )

      (LSEC:SetConfig
        "SECTIONGAP"
        120.0
      )


      ;-------------------------------------------------------
      ; Report Selected Settings
      ;-------------------------------------------------------

      (LSEC:Print
        (strcat
          "Horizontal Scale: 1:"
          (rtos hscale 2 0)
        )
      )

      (LSEC:Print
        (strcat
          "Vertical Exaggeration: "
          (rtos vexag 2 2)
        )
      )

      (LSEC:Print
        (strcat
          "Split Length: "
          (rtos splitLength 2 2)
        )
      )


      ;-------------------------------------------------------
      ; Part 7B - Prepare Final Drawing
      ;-------------------------------------------------------

      (if (LSEC:PrepareFinalDrawing)

        (progn


          ;---------------------------------------------------
          ; Start Final Split Drawing
          ;---------------------------------------------------

          (LSEC:Print
            "Generating final split longitudinal section..."
          )

          (setq result
            (LSEC:DrawAllSplitSections)
          )


          ;---------------------------------------------------
          ; Final Status
          ;---------------------------------------------------

          (if result

            (progn

              (LSEC:Print
                "========================================"
              )

              (LSEC:Print
                "FINAL DRAWING COMPLETED SUCCESSFULLY"
              )

              (LSEC:Print
                "========================================"
              )

            )

            (LSEC:Print
              "Final drawing failed."
            )

          )

        )

        (LSEC:Print
          "Final drawing preparation failed."
        )

      )

    )


    ;---------------------------------------------------------
    ; Excel File Was Not Loaded
    ;---------------------------------------------------------

    (LSEC:Print
      "Final drawing cancelled: Excel data was not loaded."
    )

  )

  (princ)

)


(LSEC:Print
  "Part 7A - Final Drawing Command Updated and Loaded."
)

;=============================================================
;
; Part 7D - Split Presentation Test
;
;=============================================================

(defun LSEC:TestFinalPresentation
       (/ splitMode splitLength sectionGap)

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "L-SECV6 Split Presentation Test"
  )

  (LSEC:Print
    "========================================"
  )

  (setq splitMode
    (LSEC:GetConfig "SPLITMODE")
  )

  (setq splitLength
    (LSEC:GetConfig "SPLITLENGTH")
  )

  (setq sectionGap
    (LSEC:GetConfig "SECTIONGAP")
  )

  (if (numberp splitMode)

    (LSEC:Print
      (strcat
        "Split Mode: "
        (rtos splitMode 2 0)
      )
    )

    (LSEC:Print
      "Split Mode: NOT SET"
    )

  )

  (if (numberp splitLength)

    (LSEC:Print
      (strcat
        "Split Length: "
        (rtos splitLength 2 2)
      )
    )

    (LSEC:Print
      "Split Length: NOT SET"
    )

  )

  (if (numberp sectionGap)

    (LSEC:Print
      (strcat
        "Section Gap: "
        (rtos sectionGap 2 2)
      )
    )

    (LSEC:Print
      "Section Gap: NOT SET"
    )

  )

  (LSEC:Print
    "Split presentation settings verified."
  )

  (LSEC:Print
    "========================================"
  )

  T

)


(LSEC:Print
  "Part 7D - Split Presentation Test Loaded."
)

;=============================================================
;
; Part 7E - Final Presentation Dispatcher
;
;=============================================================

(defun LSEC:RunFinalPresentation
       (/ splitMode)

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "L-SECV6 FINAL PRESENTATION"
  )

  (LSEC:Print
    "========================================"
  )

  ;-----------------------------------------------------------
  ; Read current split mode
  ;-----------------------------------------------------------

  (setq splitMode
    (LSEC:GetConfig "SPLITMODE")
  )

  ;-----------------------------------------------------------
  ; Split drawing
  ;-----------------------------------------------------------

  (if (= splitMode 1)

    (progn

      (LSEC:Print
        "Split mode selected."
      )

      (LSEC:DrawAllSplitSections)

    )

    ;---------------------------------------------------------
    ; Normal single-section drawing
    ;---------------------------------------------------------

    (progn

      (LSEC:Print
        "Single-section mode selected."
      )

      (LSEC:DrawCompleteSection)

    )

  )

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "Final presentation command completed."
  )

  (LSEC:Print
    "========================================"
  )

  (princ)

)


(LSEC:Print
  "Part 7E - Final Presentation Dispatcher Loaded."
)

;=============================================================
;
; Part 7F - Final User Presentation Setup
;
; Purpose:
;   Provide the final user-facing setup for L-SECV6.
;
; User enters:
;   1. Horizontal Scale
;   2. Vertical Exaggeration
;   3. Split Length
;
; Section Gap is NOT user entered.
; It is handled automatically by the split-frame engine.
;
;=============================================================


(defun LSEC:FinalPresentationSetup
       (/ hScale
          vExag
          splitLength
          splitMode
          result)

  (LSEC:Print
    "========================================"
  )

  (LSEC:Print
    "L-SECV6 FINAL PRESENTATION SETUP"
  )

  (LSEC:Print
    "========================================"
  )


  ;-----------------------------------------------------------
  ; Horizontal Scale
  ;-----------------------------------------------------------

  (setq hScale
    (getreal
      "\nEnter Horizontal Scale <1000>: "
    )
  )

  (if (not hScale)
    (setq hScale 1000.0)
  )

  (while (<= hScale 0.0)

    (princ
      "\nHorizontal Scale must be greater than zero."
    )

    (setq hScale
      (getreal
        "\nEnter Horizontal Scale <1000>: "
      )
    )

    (if (not hScale)
      (setq hScale 1000.0)
    )

  )


  ;-----------------------------------------------------------
  ; Vertical Exaggeration
  ;-----------------------------------------------------------

  (setq vExag
    (getreal
      "\nEnter Vertical Exaggeration <10>: "
    )
  )

  (if (not vExag)
    (setq vExag 10.0)
  )

  (while (<= vExag 0.0)

    (princ
      "\nVertical Exaggeration must be greater than zero."
    )

    (setq vExag
      (getreal
        "\nEnter Vertical Exaggeration <10>: "
      )
    )

    (if (not vExag)
      (setq vExag 10.0)
    )

  )


  ;-----------------------------------------------------------
  ; Split Length
  ;
  ; User enters the desired section length.
  ;
  ; Example:
  ;   200 = split every 200 m
  ;
  ; A value greater than the complete chainage will result
  ; in one section.
  ;-----------------------------------------------------------

  (setq splitLength
    (getreal
      "\nEnter Split Length <200>: "
    )
  )

  (if (not splitLength)
    (setq splitLength 200.0)
  )

  (while (<= splitLength 0.0)

    (princ
      "\nSplit Length must be greater than zero."
    )

    (setq splitLength
      (getreal
        "\nEnter Split Length <200>: "
      )
    )

    (if (not splitLength)
      (setq splitLength 200.0)
    )

  )


  ;-----------------------------------------------------------
  ; Determine split mode
  ;
  ; Split mode is automatic.
  ;
  ; If split length is greater than or equal to the complete
  ; drawing length, the drawing remains one section.
  ; Otherwise split mode is activated.
  ;-----------------------------------------------------------

  (if *LSEC-DATA*

    (progn

      (setq splitMode
        (if
          (and
            (cdr
              (assoc
                "CHAINAGE"
                (car *LSEC-DATA*)
              )
            )

            (cdr
              (assoc
                "CHAINAGE"
                (car (reverse *LSEC-DATA*))
              )
            )
          )

          (if
            (>=
              splitLength
              (-
                (cdr
                  (assoc
                    "CHAINAGE"
                    (car (reverse *LSEC-DATA*))
                  )
                )

                (cdr
                  (assoc
                    "CHAINAGE"
                    (car *LSEC-DATA*)
                  )
                )
              )
            )

            0
            1
          )

          1
        )
      )

    )

    ; No data yet.
    ; Default to split mode.
    (setq splitMode 1)

  )


  ;-----------------------------------------------------------
  ; Store configuration
  ;-----------------------------------------------------------

  (LSEC:SetConfig
    "HSCALE"
    hScale
  )

  (LSEC:SetConfig
    "VEXAG"
    vExag
  )

  (LSEC:SetConfig
    "SPLITLENGTH"
    splitLength
  )

  (LSEC:SetConfig
    "SPLITMODE"
    splitMode
  )


  ;-----------------------------------------------------------
  ; Section gap
  ;
  ; IMPORTANT:
  ; Do NOT ask the user for this.
  ;
  ; The existing section-frame calculation determines the
  ; actual frame spacing automatically.
  ;-----------------------------------------------------------

  (LSEC:Print
    "Section gap: automatic"
  )


  ;-----------------------------------------------------------
  ; Report selected settings
  ;-----------------------------------------------------------

  (LSEC:Print
    "----------------------------------------"
  )

  (LSEC:Print
    (strcat
      "Horizontal Scale: "
      (rtos hScale 2 2)
    )
  )

  (LSEC:Print
    (strcat
      "Vertical Exaggeration: "
      (rtos vExag 2 2)
    )
  )

  (LSEC:Print
    (strcat
      "Split Length: "
      (rtos splitLength 2 2)
    )
  )

  (LSEC:Print
    (strcat
      "Split Mode: "
      (if (= splitMode 1)
        "ON"
        "OFF"
      )
    )
  )

  (LSEC:Print
    "Section Gap: Automatic"
  )

  (LSEC:Print
    "----------------------------------------"
  )


  ;-----------------------------------------------------------
  ; Draw final presentation
  ;-----------------------------------------------------------

  (setq result
    (LSEC:RunFinalPresentation)
  )


  ;-----------------------------------------------------------
  ; Return result
  ;-----------------------------------------------------------

  result

)


(LSEC:Print
  "Part 7F - Final User Presentation Setup Loaded."
)

;=============================================================
;
; Part 7K - Final Excel-to-Drawing Integration
;
; Purpose:
;   Start the complete L-SECV6 workflow from one command:
;
;   LSEC
;      |
;      +-- Select Excel file
;      |
;      +-- Open workbook
;      |
;      +-- Select worksheet
;      |
;      +-- Load Excel data
;      |
;      +-- Final presentation setup
;      |
;      +-- Draw complete section
;
;=============================================================

(defun LSEC:RunFinalCommand
       (/ excelResult result)

  (princ
    "\n========================================"
  )

  (princ
    "\nL-SECV6 PROFESSIONAL FINAL DRAWING"
  )

  (princ
    "\n========================================"
  )


  ;-----------------------------------------------------------
  ; Select and load Excel data
  ;-----------------------------------------------------------

  (setq excelResult
    (LSEC:PromptAndLoadExcel)
  )


  ;-----------------------------------------------------------
  ; Excel loading failed or cancelled
  ;-----------------------------------------------------------

  (if (not excelResult)

    (progn

      (princ
        "\nL-SECV6 drawing cancelled."
      )

      nil

    )

    (progn

      ;-------------------------------------------------------
      ; Verify that data is actually available
      ;-------------------------------------------------------

      (if (not *LSEC-DATA*)

        (progn

          (princ
            "\nExcel data is empty."
          )

          (princ
            "\nL-SECV6 drawing cancelled."
          )

          nil

        )

        (progn

          ;---------------------------------------------------
          ; Run final presentation setup
          ;---------------------------------------------------

          (setq result
            (LSEC:FinalPresentationSetup)
          )


          ;---------------------------------------------------
          ; Final status
          ;---------------------------------------------------

          (if result

            (princ
              "\nL-SECV6 FINAL DRAWING COMPLETED."
            )

            (princ
              "\nL-SECV6 FINAL DRAWING FAILED."
            )

          )


          result

        )

      )

    )

  )

)


;-------------------------------------------------------------
; Main command
;-------------------------------------------------------------

(defun c:LSEC ()

  (LSEC:RunFinalCommand)

  (princ)

)


;-------------------------------------------------------------
; Alternate full command
;-------------------------------------------------------------

(defun c:LSECFINALDRAW ()

  (LSEC:RunFinalCommand)

  (princ)

)


;-------------------------------------------------------------
; Load message
;-------------------------------------------------------------

(princ
  "\nPart 7G - LSEC / LSECFINALDRAW Commands Loaded."
)

(princ)

;=============================================================
;
; Part 7H - Excel Last Location Support
;
; Purpose:
;   Remember the last Excel file location used by L-SECV6.
;
;=============================================================


;-------------------------------------------------------------
; Last Excel file
;-------------------------------------------------------------

(setq *LSEC-LAST-EXCEL-FILE* nil)


;-------------------------------------------------------------
; Get directory from complete file path
;-------------------------------------------------------------

(defun LSEC:GetExcelDirectory
       (filePath / pos)

  (if filePath

    (progn

      (setq pos
        (strlen filePath)
      )

      (while
        (and
          (> pos 0)
          (/=
            (substr filePath pos 1)
            "\\"
          )
        )

        (setq pos
          (1- pos)
        )

      )

      (if (> pos 0)

        (substr
          filePath
          1
          pos
        )

        nil

      )

    )

    nil

  )

)


;-------------------------------------------------------------
; Store last Excel file
;-------------------------------------------------------------

(defun LSEC:SetLastExcelFile
       (filePath)

  (if filePath

    (setq *LSEC-LAST-EXCEL-FILE*
      filePath
    )

  )

  *LSEC-LAST-EXCEL-FILE*

)


;-------------------------------------------------------------
; Get last Excel file
;-------------------------------------------------------------

(defun LSEC:GetLastExcelFile ()

  *LSEC-LAST-EXCEL-FILE*

)


;-------------------------------------------------------------
; Test Excel location memory
;-------------------------------------------------------------

(defun LSEC:TestExcelLocation ()

  (if *LSEC-LAST-EXCEL-FILE*

    (progn

      (LSEC:Print
        (strcat
          "Last Excel file: "
          *LSEC-LAST-EXCEL-FILE*
        )
      )

      (LSEC:Print
        (strcat
          "Last Excel folder: "
          (LSEC:GetExcelDirectory
            *LSEC-LAST-EXCEL-FILE*
          )
        )
      )

    )

    (LSEC:Print
      "No previous Excel file recorded."
    )

  )

  T

)


(LSEC:Print
  "Part 7H - Excel Last Location Support Loaded."
)

;=============================================================
;
; Part 7I - Excel Last Location Integration
;
; Purpose:
;   Integrate the remembered Excel file location with the
;   existing Excel loading workflow.
;
; IMPORTANT:
;   This part does NOT replace WorkbookOpen.
;   It provides a helper for the existing Excel file dialog.
;
;=============================================================


;-------------------------------------------------------------
; Get initial directory for Excel file dialog
;
; Returns:
;   Last Excel folder if available.
;   nil otherwise.
;-------------------------------------------------------------

(defun LSEC:GetExcelDialogDirectory
       (/ lastFile)

  (setq lastFile
    (LSEC:GetLastExcelFile)
  )

  (if lastFile

    (LSEC:GetExcelDirectory
      lastFile
    )

    nil

  )

)


;-------------------------------------------------------------
; Remember successfully opened Excel file
;-------------------------------------------------------------

(defun LSEC:RememberExcelFile
       (filePath)

  (if (and filePath
           (/= filePath ""))

    (progn

      (LSEC:SetLastExcelFile
        filePath
      )

      (LSEC:Print
        (strcat
          "Excel location remembered: "
          filePath
        )
      )

      T

    )

    nil

  )

)


;-------------------------------------------------------------
; Test remembered Excel location
;-------------------------------------------------------------

(defun LSEC:TestExcelLocationIntegration
       (/ lastFile folder)

  (setq lastFile
    (LSEC:GetLastExcelFile)
  )

  (if lastFile

    (progn

      (setq folder
        (LSEC:GetExcelDirectory
          lastFile
        )
      )

      (LSEC:Print
        "----------------------------------------"
      )

      (LSEC:Print
        (strcat
          "Remembered Excel file: "
          lastFile
        )
      )

      (if folder

        (LSEC:Print
          (strcat
            "Remembered Excel folder: "
            folder
          )
        )

      )

      (LSEC:Print
        "Excel location integration is ready."
      )

      T

    )

    (progn

      (LSEC:Print
        "No Excel file has been remembered yet."
      )

      nil

    )

  )

)


(LSEC:Print
  "Part 7I - Excel Last Location Integration Loaded."
)

;=============================================================
;
; Part 7J - Persistent Excel File Location
;
; Purpose:
;   Remember the last successfully selected Excel file
;   even after L-SECV6 is APPLOADed again.
;
;=============================================================


;-------------------------------------------------------------
; Remember last Excel file
;-------------------------------------------------------------

(defun LSEC:RememberLastExcelFile
       (filepath)

  (if
    (and
      filepath
      (= (type filepath) 'STR)
      (/= filepath "")
    )

    (progn

      ;-------------------------------------------------------
      ; Store permanently in AutoCAD environment
      ;-------------------------------------------------------

      (setenv
        "LSECV6_LastExcelFile"
        filepath
      )


      ;-------------------------------------------------------
      ; Also keep in current-session variable
      ;-------------------------------------------------------

      (setq *LSEC-LAST-EXCEL-FILE*
        filepath
      )


      (LSEC:Print
        (strcat
          "Last Excel file stored: "
          filepath
        )
      )


      filepath

    )

    nil

  )

)


;-------------------------------------------------------------
; Get last Excel file
;-------------------------------------------------------------

(defun LSEC:GetLastExcelFile
       (/ filepath)

  ;-----------------------------------------------------------
  ; First check current-session variable
  ;-----------------------------------------------------------

  (setq filepath
    *LSEC-LAST-EXCEL-FILE*
  )


  ;-----------------------------------------------------------
  ; If current-session variable is empty,
  ; recover from persistent AutoCAD environment
  ;-----------------------------------------------------------

  (if
    (or
      (not filepath)
      (= filepath "")
    )

    (setq filepath
      (getenv
        "LSECV6_LastExcelFile"
      )
    )

  )


  ;-----------------------------------------------------------
  ; Restore current-session variable
  ;-----------------------------------------------------------

  (if
    (and
      filepath
      (= (type filepath) 'STR)
      (/= filepath "")
    )

    (setq *LSEC-LAST-EXCEL-FILE*
      filepath
    )

  )


  *LSEC-LAST-EXCEL-FILE*

)


;-------------------------------------------------------------
; Get last Excel directory
;
; Supports:
;   C:/LSEC/Test.xlsx
;   C:\LSEC\Test.xlsx
;-------------------------------------------------------------

(defun LSEC:GetLastExcelDirectory
       (/ filepath pos)

  (setq filepath
    (LSEC:GetLastExcelFile)
  )


  (if filepath

    (progn

      (setq pos
        (strlen filepath)
      )


      (while
        (and
          (> pos 0)

          (/=
            (substr filepath pos 1)
            "\\"
          )

          (/=
            (substr filepath pos 1)
            "/"
          )
        )

        (setq pos
          (1- pos)
        )

      )


      (if (> pos 0)

        (substr
          filepath
          1
          pos
        )

        nil

      )

    )

    nil

  )

)


;-------------------------------------------------------------
; Select and load Excel file
;
; Uses the last remembered Excel folder as the starting
; location for the file selection dialog.
;
;-------------------------------------------------------------

(defun LSEC:PromptAndLoadExcel
       (/ startDir filepath workbook worksheet dataResult result)

  (setq result nil)

  ; Get previous Excel directory
  (setq startDir (LSEC:GetLastExcelDirectory))
  (if (or (not startDir) (= startDir ""))
    (setq startDir (getvar "DWGPREFIX"))
  )

  ; Ask user to select workbook
  (setq filepath
    (getfiled "Select Excel File" startDir "xlsx;xls" 0)
  )

  (if (not filepath)
    (LSEC:Print "Excel file selection cancelled.")
    (progn
      (LSEC:Print (strcat "Selected Excel file: " filepath))

      ; Open workbook in the L-SECV6 Excel instance
      (setq workbook (LSEC:WorkbookOpen filepath))

      (if workbook
        (progn
          ; Ask user to choose the worksheet
          (setq worksheet (LSEC:PromptWorksheetSelection))

          (if worksheet
            (progn
              ; Read worksheet data into AutoLISP memory
              (setq dataResult
                (vl-catch-all-apply 'LSEC:LoadExcelData nil)
              )

              (if (vl-catch-all-error-p dataResult)
                (progn
                  (LSEC:Print
                    (strcat
                      "Excel data loading error: "
                      (vl-catch-all-error-message dataResult)
                    )
                  )
                  (setq result nil)
                )
                (if *LSEC-DATA*
                  (progn
                    (LSEC:RememberLastExcelFile filepath)
                    (LSEC:Print
                      (strcat
                        "Excel data loaded successfully: "
                        (itoa (length *LSEC-DATA*))
                        " records."
                      )
                    )
                    (setq result T)
                  )
                  (progn
                    (LSEC:Print "Excel data loading failed or returned no records.")
                    (setq result nil)
                  )
                )
              )
            )
            (LSEC:Print "Worksheet selection cancelled or failed.")
          )
        )
        (LSEC:Print "Excel workbook could not be opened.")
      )
    )
  )

  ; Always release Excel COM references after the data has been copied
  ; into *LSEC-DATA*. Close without saving any workbook changes.
  (LSEC:ReleaseWorksheet)
  (LSEC:ExcelClose)

  result
)

(defun LSEC:TestLastExcelFile
       (/ filepath folder)

  (setq filepath
    (LSEC:GetLastExcelFile)
  )


  (setq folder
    (LSEC:GetLastExcelDirectory)
  )


  (if filepath

    (progn

      (LSEC:Print
        (strcat
          "Last Excel file: "
          filepath
        )
      )


      (if folder

        (LSEC:Print
          (strcat
            "Last Excel folder: "
            folder
          )
        )

      )


      T

    )

    (progn

      (LSEC:Print
        "No last Excel file stored."
      )

      nil

    )

  )

)


;-------------------------------------------------------------
; Load message
;-------------------------------------------------------------

(LSEC:Print
  "Part 7J - Persistent Excel File Location Loaded."
)