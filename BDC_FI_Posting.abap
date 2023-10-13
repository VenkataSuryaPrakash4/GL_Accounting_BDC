*&---------------------------------------------------------------------*
*& Report ZR_F02_BDC_POST
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zr_f02_bdc_post.
TYPES:BEGIN OF ty_excel,
        serial(5) TYPE c,
        h_txt     TYPE bktxt,
        com_code  TYPE t001-bukrs,
        doc_type  TYPE blart,
        ref_num   TYPE xblnr,
        gl_acc    TYPE hkont,
        itm_txt   TYPE sgtxt,
        pro_cnt   TYPE prctr,
        amnt      TYPE bapidoccur,
      END OF ty_excel.

DATA:lt_bdc       TYPE TABLE OF bdcdata,
     lt_excel     TYPE TABLE OF ty_excel,
     lt_msg       TYPE TABLE OF bdcmsgcoll,
     w_raw        TYPE truxs_t_text_data,
     lv_fname     TYPE rlgrap-filename,
     lv_monat     TYPE bkpf-monat,
     lv_budat     TYPE bkpf-budat,
     lv_date_Ext  TYPE char10,
     lv_amnt(132) TYPE c,
     lv_filename  TYPE ibipparms-path.

*1. Selction screen to read the file from local computer
PARAMETERS:p_file TYPE dynpread-fieldname.

*  1.1 Enable F4 help to input field.
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.

* 1.2 Display pop up to choose the file from local computer.
  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = syst-cprog
      dynpro_number = syst-dynnr
      field_name    = 'P_FILE'
    IMPORTING
      file_name     = lv_filename.
  IF sy-subrc = 0.
    p_file = lv_filename.
    lv_fname = lv_filename.
  ENDIF.

START-OF-SELECTION.
*2 Convert the excel file to internal tables
  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_field_seperator    = 'X'
      i_line_header        = 'X'
      i_tab_raw_data       = w_raw
      i_filename           = lv_fname
    TABLES
      i_tab_converted_data = lt_excel
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

  IF lt_excel IS NOT INITIAL.
    SELECT bukrs,waers
      FROM t001
      INTO TABLE @DATA(lt_t001)
      FOR ALL ENTRIES IN @lt_excel
      WHERE bukrs = @lt_excel-com_code.
    IF sy-subrc = 0.
      SORT lt_t001 BY bukrs.
    ENDIF.
  ENDIF.
  LOOP AT lt_excel INTO DATA(wa_excel_tmp).
    DATA(wa_excel) = wa_excel_tmp.
    AT NEW serial.
      lv_budat = sy-datum.
      CALL FUNCTION 'FI_PERIOD_DETERMINE'
        EXPORTING
          i_budat        = lv_budat
          i_bukrs        = wa_excel-com_code
        IMPORTING
          e_monat        = lv_monat
        EXCEPTIONS
          fiscal_year    = 1
          period         = 2
          period_version = 3
          posting_period = 4
          special_period = 5
          version        = 6
          posting_date   = 7
          OTHERS         = 8.

      CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
        EXPORTING
          date_internal            = sy-datum
        IMPORTING
          date_external            = lv_date_Ext
        EXCEPTIONS
          date_internal_is_invalid = 1
          OTHERS                   = 2.

      lv_amnt = wa_excel-amnt.
      CONDENSE lv_amnt NO-GAPS.

      REPLACE ALL OCCURRENCES OF '.' IN lv_amnt WITH ','.

      DATA(wa_t001) = lt_t001[ bukrs = wa_excel-com_code ].
      lt_bdc = VALUE #( ( program = 'SAPMF05A' dynpro = '0100' dynbegin = 'X'  )
                        ( fnam = 'BDC_CURSOR'  fval = 'RF05A-NEWKO' )
                        ( fnam = 'BDC_OKCODE' fval = '/00' )
                        ( fnam = 'BKPF-BLDAT' fval = lv_date_Ext )
                        ( fnam = 'BKPF-BUKRS' fval = wa_excel-com_code )
                        ( fnam = 'BKPF-BUDAT' fval = lv_date_Ext )
                        ( fnam = 'BKPF-MONAT' fval = lv_monat )
                        ( fnam = 'BKPF-WAERS' fval = wa_t001-waers )
                        ( fnam = 'FS006-DOCID' fval = '*' )
                        ( fnam = 'RF05A-NEWBS' fval = '40' )
                        ( fnam = 'RF05A-NEWKO' fval = wa_excel-gl_acc )

                        ( program = 'SAPMF05A' dynpro = '0300' dynbegin = 'X'  )
                        ( fnam = 'BDC_CURSOR'  fval = 'BSEG-WRBTR' )
                        ( fnam = 'BDC_OKCODE' fval = '/00' )
                        ( fnam = 'BSEG-WRBTR' fval = lv_amnt )
                        ( fnam = 'DKACB-FMORE' fval = 'X' )

                        ( program = 'SAPLKACB' dynpro = '0002' dynbegin = 'X'  )
                        ( fnam = 'BDC_CURSOR'  fval = 'COBL-PRCTR' )
                        ( fnam = 'BDC_OKCODE' fval = '=ENTE' )
                        ( fnam = 'COBL-PRCTR' fval = wa_excel-pro_cnt ) ).
    ENDAT.

*Credit details
    APPEND VALUE #( program = 'SAPMF05A' dynpro = '0300' dynbegin = 'X'  ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_CURSOR'  fval = 'RF05A-NEWKO' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_OKCODE' fval = '/00' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BSEG-WRBTR' fval = lv_amnt ) TO lt_bdc.
    APPEND VALUE #( fnam = 'RF05A-NEWBS' fval = '50' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'RF05A-NEWKO' fval = wa_excel-gl_acc ) TO lt_bdc.
    APPEND VALUE #( fnam = 'DKACB-FMORE' fval = 'X' ) TO lt_bdc.

    APPEND VALUE #( program = 'SAPLKACB' dynpro = '0002' dynbegin = 'X'  ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_CURSOR'  fval = 'COBL-PRCTR' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_OKCODE' fval = '=ENTE' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'COBL-PRCTR' fval = wa_excel-pro_cnt ) TO lt_bdc.

    APPEND VALUE #( program = 'SAPMF05A' dynpro = '0300' dynbegin = 'X'  ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_CURSOR'  fval = 'BSEG-WRBTR' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_OKCODE' fval = '=BU' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BSEG-WRBTR' fval = lv_amnt ) TO lt_bdc.
    APPEND VALUE #( fnam = 'DKACB-FMORE' fval = 'X' ) TO lt_bdc.

    APPEND VALUE #( program = 'SAPLKACB' dynpro = '0002' dynbegin = 'X'  ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_CURSOR'  fval = 'COBL-PRCTR' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'BDC_OKCODE' fval = '=ENTE' ) TO lt_bdc.
    APPEND VALUE #( fnam = 'COBL-PRCTR' fval = wa_excel-pro_cnt ) TO lt_bdc.
    AT END OF serial.
      CALL TRANSACTION 'F-02' WITH AUTHORITY-CHECK
                              USING lt_bdc
                              MODE 'A' "A - ALL screens / E - Error screens / N - No screens
                              UPDATE 'S' "S - syncrnouns / A - Asyncronous
                              MESSAGES INTO lt_msg.
      CLEAR: lt_bdc.
    ENDAT.
    CLEAR:wa_excel,wa_excel_tmp,lv_monat.
  ENDLOOP.
