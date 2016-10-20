Public Class DataPathDefinition
    Private Const C_ROOT_DATA_PATH As String = "C:\業務管理ソフトData"

    'Root Folder以下のFoloder
    Private Const C_PRODUCT_FOLDER_NAME As String = "商品情報"
    Private Const C_TRADE_FOLDER_NAME As String = "取引情報"
    Private Const C_TEMPLATE_FOLDER_NAME As String = "Template情報"

    'Template情報 Folder内のFile
    Private Const C_TEMPLATE_ZAIKO_FILE_NAME As String = "Template_在庫.xlsx"
    Private Const C_TEMPLATE_SYUKA_FILE_NAME As String = "Template_出荷.xlsx"
    Private Const C_TEMPLATE_NYUKA_FILE_NAME As String = "Template_入荷.xlsx"

    '取引情報 以下のFolder
    Private Const C_SHIIRE_FOLDER_NAME As String = "仕入れ情報"
    Private Const C_CUSTOMER_FOLDER_NAME As String = "お客様"


    Public Shared ReadOnly Property GetRootDataPath() As String
        Get
            Return C_ROOT_DATA_PATH
        End Get
    End Property

#Region "Root以下のFolder"
    ''' <summary>
    ''' 商品情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetProductDataPath() As String
        Get
            Return C_ROOT_DATA_PATH + "\" + C_PRODUCT_FOLDER_NAME
        End Get
    End Property

    ''' <summary>
    ''' Template情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetTemplateDataPath() As String
        Get
            Return C_ROOT_DATA_PATH + "\" + C_TEMPLATE_FOLDER_NAME
        End Get
    End Property

    ''' <summary>
    ''' 取引情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetTradeDataPath() As String
        Get
            Return C_ROOT_DATA_PATH + "\" + C_TRADE_FOLDER_NAME
        End Get
    End Property

#End Region

#Region "Template情報Folder内のFile"
    ''' <summary>
    ''' Template在庫情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetTemplateZaikoPath() As String
        Get
            Return GetTemplateDataPath() + "\" + C_TEMPLATE_ZAIKO_FILE_NAME
        End Get
    End Property

    ''' <summary>
    ''' Template出荷情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetTemplateSyukaPath() As String
        Get
            Return GetTemplateDataPath() + "\" + C_TEMPLATE_SYUKA_FILE_NAME
        End Get
    End Property

    ''' <summary>
    ''' Template入荷情報
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetTemplateNyukaPath() As String
        Get
            Return GetTemplateDataPath() + "\" + C_TEMPLATE_NYUKA_FILE_NAME
        End Get
    End Property

#End Region

#Region "取引情報以下のFolder"
    ''' <summary>
    ''' 仕入れ情報のFolderパス
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetShiireDataPath() As String
        Get
            Return GetTradeDataPath() + "\" + C_SHIIRE_FOLDER_NAME
        End Get
    End Property

    ''' <summary>
    ''' お客様情報のFolder
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property GetCustomerDataPath() As String
        Get
            Return GetTradeDataPath() + "\" + C_CUSTOMER_FOLDER_NAME
        End Get
    End Property

#End Region

End Class
