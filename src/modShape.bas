Attribute VB_Name = "modShape"
Option Explicit

''' �}�`�̎��
Public Type ShapeCategory
    sharpType As MsoShapeType
    name As String
    hasText As Boolean
    hasImage As Boolean
End Type

''' �}�`�̎�ނ�������
Public Function ShapeCategoriesInitialize() As ShapeCategory()
    ReDim ShapeCategoriesInitialize(30) As ShapeCategory
    Call ShapeCategorySet(ShapeCategoriesInitialize, 0, mso3DModel,             "3D ���f��",                    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 1, msoAutoShape,           "�I�[�g�V�F�C�v",               True, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 2, msoCallout,             "�����o��",                     True, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 3, msoCanvas,              "�L�����o�X",                   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 4, msoChart,               "�O���t",                       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 5, msoComment,             "�R�����g",                     True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 6, msoContentApp,          "�R���e���c Office �A�h�C��",   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 7, msoDiagram,             "�}",                           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 8, msoEmbeddedOLEObject,   "���ߍ��� OLE �I�u�W�F�N�g",    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 9, msoFormControl,         "�t�H�[�� �R���g���[��",        False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 10, msoFreeform,           "�t���[�t�H�[��",               False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 11, msoGraphic,            "�O���t�B�b�N",                 False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 12, msoGroup,              "Group",                        False, False)
   'Call ShapeCategorySet(ShapeCategoriesInitialize, 12, msoIgxGraphic,         "SmartArt �O���t�B�b�N",        False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 13, msoInk,                "�C���N",                       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 14, msoInkComment,         "�C���N �R�����g",              False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 15, msoLine,               "Line",                         False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 16, msoLinked3DModel,      "�����N���ꂽ 3D ���f��",       False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 17, msoLinkedGraphic,      "�����N���ꂽ�O���t�B�b�N",     False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 18, msoLinkedOLEObject,    "�����N OLE �I�u�W�F�N�g",      False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 19, msoLinkedPicture,      "�����N�摜",                   False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 20, msoMedia,              "���f�B�A",                     False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 21, msoOLEControlObject,   "OLE �R���g���[�� �I�u�W�F�N�g",    False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 22, msoPicture,            "�摜",                             False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 23, msoPlaceholder,        "�v���[�X�z���_�[",                 False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 24, msoScriptAnchor,       "�X�N���v�g �A���J�[",              False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 25, msoShapeTypeMixed,     "�}�`�̎�ނ̑g�ݍ��킹",           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 26, msoSlicer,             "Slicer",                           False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 27, msoTable,              "�e�[�u��",                         False, True)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 28, msoTextBox,            "�e�L�X�g �{�b�N�X",                True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 29, msoTextEffect,         "�e�L�X�g����",                     True, False)
    Call ShapeCategorySet(ShapeCategoriesInitialize, 30, msoWebVideo,           "Web �r�f�I",                       False, True)
End Function

''' �}�`�̎�ނ�ǉ�
Private Sub ShapeCategorySet(ByRef categories() As ShapeCategory, index As Integer, sharpType As MsoShapeType, name As String, hasText As Boolean, hasImage As Boolean)
    categories(index).sharpType = sharpType
    categories(index).name = name
    categories(index).hasText = hasText
    categories(index).hasImage = hasImage
End Sub

''' �}�`�̎�ނ�T��
Public Function ShapeCategoriesFind(ByRef categories() As ShapeCategory, sharpType As MsoShapeType) As ShapeCategory
    ShapeCategoriesFind.sharpType = -1
    ShapeCategoriesFind.name = "[�s��]"
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
