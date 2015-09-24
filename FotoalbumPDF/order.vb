Public Class order

    Private _id As String
    Private _status As String
    Private _order_id As String
    Private _product_id As String
    Private _user_product_id As String
    Private _name As String
    Private _pages_xml As String
    Private _textflow_xml As String
    Private _textlines_xml As String
    Private _photo_xml As String
    Private _color_xml As String
    Private _platform As String

    Public Sub New()

    End Sub

    Property id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            Me._id = value
        End Set
    End Property

    Property status() As String
        Get
            Return _status
        End Get
        Set(ByVal value As String)
            Me._status = value
        End Set
    End Property

    Property order_id() As String
        Get
            Return _order_id
        End Get
        Set(ByVal value As String)
            Me._order_id = value
        End Set
    End Property

    Property product_id() As String
        Get
            Return _product_id
        End Get
        Set(ByVal value As String)
            Me._product_id = value
        End Set
    End Property

    Property user_product_id() As String
        Get
            Return _user_product_id
        End Get
        Set(ByVal value As String)
            Me._user_product_id = value
        End Set
    End Property

    Property name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            Me._name = value
        End Set
    End Property

    Property pages_xml() As String
        Get
            Return _pages_xml
        End Get
        Set(ByVal value As String)
            Me._pages_xml = value
        End Set
    End Property

    Property textflow_xml() As String
        Get
            Return _textflow_xml
        End Get
        Set(ByVal value As String)
            Me._textflow_xml = value
        End Set
    End Property

    Property textlines_xml() As String
        Get
            Return _textlines_xml
        End Get
        Set(ByVal value As String)
            Me._textlines_xml = value
        End Set
    End Property

    Property photo_xml() As String
        Get
            Return _photo_xml
        End Get
        Set(ByVal value As String)
            Me._photo_xml = value
        End Set
    End Property

    Property color_xml() As String
        Get
            Return _color_xml
        End Get
        Set(ByVal value As String)
            Me._color_xml = value
        End Set
    End Property

    Property platform() As String
        Get
            Return _platform
        End Get
        Set(ByVal value As String)
            Me._platform = value
        End Set
    End Property

End Class
