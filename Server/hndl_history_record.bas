Attribute VB_Name = "hndl_history_record"
Option Explicit


Public Function is_relevant(rg_record As Range) As Boolean
    Select Case rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_STARTED).Value
        Case db_transaction_type.STR_TRANSACTION_TYPE_ASN_GR, _
                db_transaction_type.STR_TRANSACTION_TYPE_PO_GR, _
                db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_GI, _
                db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_GR, _
                db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_PICK, _
                db_transaction_type.STR_TRANSACTION_TYPE_TASK_LIST_CREATE, _
                db_transaction_type.STR_TRANSACTION_TYPE_HU_MOVE, _
                db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU
            is_relevant = True
        Case Else
            is_relevant = False
    End Select
End Function
