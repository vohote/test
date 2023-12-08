class ZCL_EXCEL_STYLE_CONDITIONAL definition
  public
  final
  create public .

public section.

  constants C_CFVO_TYPE_FORMULA type ZEXCEL_CONDITIONAL_TYPE value 'formula' ##NO_TEXT.
  constants C_CFVO_TYPE_MAX type ZEXCEL_CONDITIONAL_TYPE value 'max' ##NO_TEXT.
  constants C_CFVO_TYPE_MIN type ZEXCEL_CONDITIONAL_TYPE value 'min' ##NO_TEXT.
  constants C_CFVO_TYPE_NUMBER type ZEXCEL_CONDITIONAL_TYPE value 'num' ##NO_TEXT.
  constants C_CFVO_TYPE_PERCENT type ZEXCEL_CONDITIONAL_TYPE value 'percent' ##NO_TEXT.
  constants C_CFVO_TYPE_PERCENTILE type ZEXCEL_CONDITIONAL_TYPE value 'percentile' ##NO_TEXT.
  constants C_ICONSET_3ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '3Arrows' ##NO_TEXT.
  constants C_ICONSET_3ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '3ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_3FLAGS type ZEXCEL_CONDITION_RULE_ICONSET value '3Flags' ##NO_TEXT.
  constants C_ICONSET_3SIGNS type ZEXCEL_CONDITION_RULE_ICONSET value '3Signs' ##NO_TEXT.
  constants C_ICONSET_3SYMBOLS type ZEXCEL_CONDITION_RULE_ICONSET value '3Symbols' ##NO_TEXT.
  constants C_ICONSET_3SYMBOLS2 type ZEXCEL_CONDITION_RULE_ICONSET value '3Symbols2' ##NO_TEXT.
  constants C_ICONSET_3TRAFFICLIGHTS type ZEXCEL_CONDITION_RULE_ICONSET value '' ##NO_TEXT.
  constants C_ICONSET_3TRAFFICLIGHTS2 type ZEXCEL_CONDITION_RULE_ICONSET value '3TrafficLights2' ##NO_TEXT.
  constants C_ICONSET_4ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '4Arrows' ##NO_TEXT.
  constants C_ICONSET_4ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '4ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_4RATING type ZEXCEL_CONDITION_RULE_ICONSET value '4Rating' ##NO_TEXT.
  constants C_ICONSET_4REDTOBLACK type ZEXCEL_CONDITION_RULE_ICONSET value '4RedToBlack' ##NO_TEXT.
  constants C_ICONSET_4TRAFFICLIGHTS type ZEXCEL_CONDITION_RULE_ICONSET value '4TrafficLights' ##NO_TEXT.
  constants C_ICONSET_5ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '5Arrows' ##NO_TEXT.
  constants C_ICONSET_5ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '5ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_5QUARTERS type ZEXCEL_CONDITION_RULE_ICONSET value '5Quarters' ##NO_TEXT.
  constants C_ICONSET_5RATING type ZEXCEL_CONDITION_RULE_ICONSET value '5Rating' ##NO_TEXT.
  constants C_OPERATOR_BEGINSWITH type ZEXCEL_CONDITION_OPERATOR value 'beginsWith' ##NO_TEXT.
  constants C_OPERATOR_BETWEEN type ZEXCEL_CONDITION_OPERATOR value 'between' ##NO_TEXT.
  constants C_OPERATOR_CONTAINSTEXT type ZEXCEL_CONDITION_OPERATOR value 'containsText' ##NO_TEXT.
  constants C_OPERATOR_ENDSWITH type ZEXCEL_CONDITION_OPERATOR value 'endsWith' ##NO_TEXT.
  constants C_OPERATOR_EQUAL type ZEXCEL_CONDITION_OPERATOR value 'equal' ##NO_TEXT.
  constants C_OPERATOR_GREATERTHAN type ZEXCEL_CONDITION_OPERATOR value 'greaterThan' ##NO_TEXT.
  constants C_OPERATOR_GREATERTHANOREQUAL type ZEXCEL_CONDITION_OPERATOR value 'greaterThanOrEqual' ##NO_TEXT.
  constants C_OPERATOR_LESSTHAN type ZEXCEL_CONDITION_OPERATOR value 'lessThan' ##NO_TEXT.
  constants C_OPERATOR_LESSTHANOREQUAL type ZEXCEL_CONDITION_OPERATOR value 'lessThanOrEqual' ##NO_TEXT.
  constants C_OPERATOR_NONE type ZEXCEL_CONDITION_OPERATOR value '' ##NO_TEXT.
  constants C_OPERATOR_NOTCONTAINS type ZEXCEL_CONDITION_OPERATOR value 'notContains' ##NO_TEXT.
  constants C_OPERATOR_NOTEQUAL type ZEXCEL_CONDITION_OPERATOR value 'notEqual' ##NO_TEXT.
  constants C_RULE_CELLIS type ZEXCEL_CONDITION_RULE value 'cellIs' ##NO_TEXT.
  constants C_RULE_CONTAINSTEXT type ZEXCEL_CONDITION_RULE value 'containsText' ##NO_TEXT.
  constants C_RULE_DATABAR type ZEXCEL_CONDITION_RULE value 'dataBar' ##NO_TEXT.
  constants C_RULE_EXPRESSION type ZEXCEL_CONDITION_RULE value 'expression' ##NO_TEXT.
  constants C_RULE_ICONSET type ZEXCEL_CONDITION_RULE value 'iconSet' ##NO_TEXT.
  constants C_RULE_COLORSCALE type ZEXCEL_CONDITION_RULE value 'colorScale' ##NO_TEXT.
  constants C_RULE_NONE type ZEXCEL_CONDITION_RULE value 'none' ##NO_TEXT.
  constants C_RULE_TOP10 type ZEXCEL_CONDITION_RULE value 'top10' ##NO_TEXT.
  constants C_RULE_ABOVE_AVERAGE type ZEXCEL_CONDITION_RULE value 'aboveAverage' ##NO_TEXT.
  constants C_SHOWVALUE_FALSE type ZEXCEL_CONDITIONAL_SHOW_VALUE value 0 ##NO_TEXT.
  constants C_SHOWVALUE_TRUE type ZEXCEL_CONDITIONAL_SHOW_VALUE value 1 ##NO_TEXT.
  data MODE_CELLIS type ZEXCEL_CONDITIONAL_CELLIS .
  data MODE_COLORSCALE type ZEXCEL_CONDITIONAL_COLORSCALE .
  data MODE_DATABAR type ZEXCEL_CONDITIONAL_DATABAR .
  data MODE_EXPRESSION type ZEXCEL_CONDITIONAL_EXPRESSION .
  data MODE_ICONSET type ZEXCEL_CONDITIONAL_ICONSET .
  data MODE_TOP10 type ZEXCEL_CONDITIONAL_TOP10 .
  data MODE_ABOVE_AVERAGE type ZEXCEL_CONDITIONAL_ABOVE_AVG .
  data PRIORITY type ZEXCEL_STYLE_PRIORITY value 1 ##NO_TEXT.
  data RULE type ZEXCEL_CONDITION_RULE .

  methods CONSTRUCTOR .
  methods GET_DIMENSION_RANGE
    returning
      value(EP_DIMENSION_RANGE) type STRING .
  methods SET_RANGE
    importing
      !IP_START_ROW type ZEXCEL_CELL_ROW
      !IP_START_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_STOP_ROW type ZEXCEL_CELL_ROW
      !IP_STOP_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  methods ADD_RANGE
    importing
      !IP_START_ROW type ZEXCEL_CELL_ROW
      !IP_START_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_STOP_ROW type ZEXCEL_CELL_ROW
      !IP_STOP_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  class-methods FACTORY_COND_STYLE_ICONSET
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IV_ICON_TYPE type ZEXCEL_CONDITION_RULE_ICONSET default C_ICONSET_3TRAFFICLIGHTS2
      !IV_CFVO1_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO1_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO2_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO2_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO3_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO3_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO4_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO4_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO5_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO5_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_SHOWVALUE type ZEXCEL_CONDITIONAL_SHOW_VALUE default ZCL_EXCEL_STYLE_CONDITIONAL=>C_SHOWVALUE_TRUE
    returning
      value(RV_STYLE_CONDITIONAL) type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
protected section.
*"* protected components of class ZCL_EXCEL_STYLE_CONDITIONAL
*"* do not include other source files here!!!
private section.

  data MV_RULE_RANGE type STRING .
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_CONDITIONAL IMPLEMENTATION.


METHOD add_range.
  DATA: lv_column    TYPE zexcel_cell_column,
        lv_row_alpha TYPE string,
        lv_col_alpha TYPE string,
        lv_coords1   TYPE string,
        lv_coords2   TYPE string.


  lv_column = zcl_excel_common=>convert_column2int( ip_start_column ).
*  me->mv_cell_data-cell_row     = 1.
*  me->mv_cell_data-cell_column  = lv_column.
*
  lv_col_alpha = ip_start_column.
  lv_row_alpha = ip_start_row.
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords1.

  IF ip_stop_column IS NOT INITIAL.
    lv_column = zcl_excel_common=>convert_column2int( ip_stop_column ).
  ELSE.
    lv_column = zcl_excel_common=>convert_column2int( ip_start_column ).
  ENDIF.

  IF ip_stop_row IS NOT INITIAL. " If we don't get explicitly a stop column use start column
    lv_row_alpha = ip_stop_row.
  ELSE.
    lv_row_alpha = ip_start_row.
  ENDIF.
  IF ip_stop_column IS NOT INITIAL. " If we don't get explicitly a stop column use start column
    lv_col_alpha = ip_stop_column.
  ELSE.
    lv_col_alpha = ip_start_column.
  ENDIF.
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords2.
  IF lv_coords2 IS NOT INITIAL AND lv_coords2 <> lv_coords1.
    CONCATENATE me->mv_rule_range ` ` lv_coords1 ':' lv_coords2 INTO me->mv_rule_range.
  ELSE.
    CONCATENATE me->mv_rule_range ` ` lv_coords1  INTO me->mv_rule_range.
  ENDIF.
  SHIFT me->mv_rule_range LEFT DELETING LEADING space.

ENDMETHOD.


METHOD constructor.

  DATA: ls_iconset TYPE zexcel_conditional_iconset.
  ls_iconset-iconset     = zcl_excel_style_conditional=>c_iconset_3trafficlights.
  ls_iconset-cfvo1_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo1_value = '0'.
  ls_iconset-cfvo2_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo2_value = '20'.
  ls_iconset-cfvo3_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo3_value = '40'.
  ls_iconset-cfvo4_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo4_value = '60'.
  ls_iconset-cfvo5_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo5_value = '80'.


  me->rule          = zcl_excel_style_conditional=>c_rule_none.
*  me->iconset->operator    = zcl_excel_style_conditional=>c_operator_none.
  me->mode_iconset  = ls_iconset.
  me->priority      = 1.

* inizialize dimension range
  me->MV_RULE_RANGE     = 'A1'.
ENDMETHOD.


METHOD factory_cond_style_iconset.

*--------------------------------------------------------------------*
* Work in progress
* Missing:  LE or LT may be specified --> extend structure ZEXCEL_CONDITIONAL_ICONSET to hold this information as well
*--------------------------------------------------------------------*

*  DATA: lv_needed_values TYPE i.
*  CASE icon_type.
*
*    WHEN 'C_ICONSET_3ARROWS'
*      OR 'C_ICONSET_3ARROWSGRAY'
*      OR 'C_ICONSET_3FLAGS'
*      OR 'C_ICONSET_3SIGNS'
*      OR 'C_ICONSET_3SYMBOLS'
*      OR 'C_ICONSET_3SYMBOLS2'
*      OR 'C_ICONSET_3TRAFFICLIGHTS'
*      OR 'C_ICONSET_3TRAFFICLIGHTS2'.
*      lv_needed_values = 3.
*
*    WHEN 'C_ICONSET_4ARROWS'
*      OR 'C_ICONSET_4ARROWSGRAY'
*      OR 'C_ICONSET_4RATING'
*      OR 'C_ICONSET_4REDTOBLACK'
*      OR 'C_ICONSET_4TRAFFICLIGHTS'.
*      lv_needed_values = 4.
*
*    WHEN 'C_ICONSET_5ARROWS'
*      OR 'C_ICONSET_5ARROWSGRAY'
*      OR 'C_ICONSET_5QUARTERS'
*      OR 'C_ICONSET_5RATING'.
*      lv_needed_values = 5.
*
*    WHEN OTHERS.
*      RETURN.
*  ENDCASE.

ENDMETHOD.


METHOD get_dimension_range.

  ep_dimension_range = me->mv_rule_range.

ENDMETHOD.


METHOD set_range.

  CLEAR: me->mv_rule_range.

  me->add_range( ip_start_row    = ip_start_row
                 ip_start_column = ip_start_column
                 ip_stop_row     = ip_stop_row
                 ip_stop_column  = ip_stop_column ).

ENDMETHOD.
ENDCLASS.
