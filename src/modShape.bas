Attribute VB_Name = "modShape"
Option Explicit

''' 図形の種類
Public Type ShapeCategory
    sharpType As MsoShapeType
    name As String
    hasText As Boolean
    hasImage As Boolean
End Type

''' 図形の種類を初期化
Public Function ShapeCategoriesInitialize() As ShapeCategory()
    ReDim ShapeCategoriesInitialize(30) As ShapeCategory
    Call ShapeCategorySet(ShapeCategoriesInitialize, 0, mso3DModel,             "3D モデル",                    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 1, msoAutoShape,           "オートシェイプ",               True, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 2, msoCallout,             "吹き出し",                     True, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 3, msoCanvas,              "キャンバス",                   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 4, msoChart,               "グラフ",                       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 5, msoComment,             "コメント",                     True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 6, msoContentApp,          "コンテンツ Office アドイン",   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 7, msoDiagram,             "図",                           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 8, msoEmbeddedOLEObject,   "埋め込み OLE オブジェクト",    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 9, msoFormControl,         "フォーム コントロール",        False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 10, msoFreeform,           "フリーフォーム",               False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 11, msoGraphic,            "グラフィック",                 False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 12, msoGroup,              "Group",                        False, False)
   'Call ShapeCategorySet(ShapeCategoriesInitialize, 12, msoIgxGraphic,         "SmartArt グラフィック",        False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 13, msoInk,                "インク",                       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 14, msoInkComment,         "インク コメント",              False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 15, msoLine,               "Line",                         False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 16, msoLinked3DModel,      "リンクされた 3D モデル",       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 17, msoLinkedGraphic,      "リンクされたグラフィック",     False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 18, msoLinkedOLEObject,    "リンク OLE オブジェクト",      False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 19, msoLinkedPicture,      "リンク画像",                   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 20, msoMedia,              "メディア",                     False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 21, msoOLEControlObject,   "OLE コントロール オブジェクト",    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 22, msoPicture,            "画像",                             False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 23, msoPlaceholder,        "プレースホルダー",                 False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 24, msoScriptAnchor,       "スクリプト アンカー",              False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 25, msoShapeTypeMixed,     "図形の種類の組み合わせ",           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 26, msoSlicer,             "Slicer",                           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 27, msoTable,              "テーブル",                         False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 28, msoTextBox,            "テキスト ボックス",                True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 29, msoTextEffect,         "テキスト効果",                     True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 30, msoWebVideo,           "Web ビデオ",                       False, True)
End Function

''' 図形の種類を追加
Private Sub ShapeCategorySet(ByRef categories() As ShapeCategory, index As Integer, sharpType As MsoShapeType, name As String, hasText As Boolean, hasImage As Boolean)
    categories(index).sharpType = sharpType
    categories(index).name = name
    categories(index).hasText = hasText
    categories(index).hasImage = hasImage
End Sub

''' 図形の種類を探す
Public Function ShapeCategoriesFind(ByRef categories() As ShapeCategory, sharpType As MsoShapeType) As ShapeCategory
    ShapeCategoriesFind.sharpType = -1
    ShapeCategoriesFind.name = "[不明]"
    ShapeCategoriesFind.hasText = False
    ShapeCategoriesFind.hasImage = False
    Dim index As Integer
    For index = LBound(categories) To UBound(categories)
        If categories(index).sharpType = sharpType Then
            ShapeCategoriesFind = categories(index)
            Exit For
        End If
    Next index
End Function
