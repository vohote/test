class ZCL_EXCEL_CONSTANT definition
  public
  final
  create public .

public section.

  constants:
    BEGIN OF gc_screen_mode.
    CONSTANTS : display TYPE c LENGTH 1 VALUE 'D',
                create  TYPE c LENGTH 1 VALUE 'N',
                change  TYPE c LENGTH 1 VALUE 'C'.
    CONSTANTS END OF gc_screen_mode .
  constants:
    BEGIN OF cs_bpg1,
        z1100 TYPE bu_group VALUE '1100',
        z1200 TYPE bu_group VALUE '1200',
        z1300 TYPE bu_group VALUE '1300',
        z1400 TYPE bu_group VALUE '1400',
        z1500 TYPE bu_group VALUE '1500',
        z1600 TYPE bu_group VALUE '1600',
        z2100 TYPE bu_group VALUE '2100',
        z2200 TYPE bu_group VALUE '2200',
        z2300 TYPE bu_group VALUE '2300',
        z2400 TYPE bu_group VALUE '2400',
        z2500 TYPE bu_group VALUE '2500',
        z2600 TYPE bu_group VALUE '2600',
        hld1  TYPE bu_group VALUE 'HLD1',
        hld2  TYPE bu_group VALUE 'HLD2',
        hld3  TYPE bu_group VALUE 'HLD3',
        hld4  TYPE bu_group VALUE 'HLD4',
        hld5  TYPE bu_group VALUE 'HLD5',
        hld6  TYPE bu_group VALUE 'HLD6',
        hlk1  TYPE bu_group VALUE 'HLK1',
        hlk2  TYPE bu_group VALUE 'HLK2',
        hlk3  TYPE bu_group VALUE 'HLK3',
        hlk4  TYPE bu_group VALUE 'HLK4',
        hlk5  TYPE bu_group VALUE 'HLK5',
        hlk6  TYPE bu_group VALUE 'HLK6',
        hlk7  TYPE bu_group VALUE 'HLK7',
      END OF cs_bpg1 .
  constants:
    BEGIN OF cs_bpg2,
        c_kunnr TYPE bu_role_screen VALUE 'HLSD00X', "'G10000X',
        c_lifnr TYPE bu_role_screen VALUE 'HLSK00X', "'G20000X',
      END OF cs_bpg2 .
  class-data CV_BUKRS type BUKRS value 'HL00' ##NO_TEXT.
  class-data CV_KOKRS type KOKRS value 'HL00' ##NO_TEXT.
  class-data CV_WERKS type WERKS_D value '1000' ##NO_TEXT.
  class-data CV_VKORG type VKORG value '1000' ##NO_TEXT.
  class-data CV_EKORG type EKORG value '1000' ##NO_TEXT.
  class-data CV_KTOPL type KTOPL value 'HL00' ##NO_TEXT.
  class-data CV_BUPLA type BUPLA value 'HLS1' ##NO_TEXT.
  class-data CS_SFPOUTPUTPARAMS type SFPOUTPUTPARAMS .
  class-data CV_EMAIL type CHAR50 value '' ##NO_TEXT.
  class-data CV_SYSID_P type SYSID value 'HEP' ##NO_TEXT.
  class-data CV_SYSID_Q type SYSID value 'HEQ' ##NO_TEXT.
  class-data CE1 type STRING value 'CE1HL00' ##NO_TEXT.

  class-methods GET_SFPOUTPUTPARAMS
    returning
      value(RS_SFPOUTPUTPARAMS) type SFPOUTPUTPARAMS .
protected section.
PRIVATE SECTION.

  TYPES ty_sfpoutputparams TYPE sfpoutputparams .
ENDCLASS.



CLASS ZCL_EXCEL_CONSTANT IMPLEMENTATION.


  METHOD GET_SFPOUTPUTPARAMS.

    cs_sfpoutputparams-reqimm = abap_on.

    rs_sfpoutputparams = cs_sfpoutputparams.

  ENDMETHOD.
ENDCLASS.
