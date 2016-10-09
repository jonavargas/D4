Imports Microsoft.Office.Interop.Word 'control de office
Imports System.IO 'sistema de archivos
Imports Microsoft.Office.Interop

Public Class D4

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click
        Dim MSWord As New Word.Application
        Dim Documento As Word.Document

        MsgBox("El TDR se guardará en : C:\Users\Jonathan\Desktop\salida.doc")
        FileCopy("C:\documento.doc", "C:\Users\Jonathan\Desktop\salida.doc")
        Documento = MSWord.Documents.Open("C:\Users\Jonathan\Desktop\salida.doc")

        Documento.Bookmarks.Item("cliente").Range.Text = " " + txtCliente.Text
        Documento.Bookmarks.Item("modulo").Range.Text = " " + txtModulo.Text
        Documento.Bookmarks.Item("paso").Range.Text = " " + txtPaso.Text
        Documento.Bookmarks.Item("MO").Range.Text = " " + txtModuloOblicuo.Text
        Documento.Bookmarks.Item("numFresa").Range.Text = " " + txtNumFresa.Text
        Documento.Bookmarks.Item("numDientes").Range.Text = " " + txtNumDientes.Text
        Documento.Bookmarks.Item("TEngranaje").Range.Text = " " + txtTipoEngranaje.Text
        Documento.Bookmarks.Item("DExterior").Range.Text = " " + txtDiametroExterior.Text
        Documento.Bookmarks.Item("DPrimitivo").Range.Text = " " + txtDiametroPrimitivo.Text
        Documento.Bookmarks.Item("DFondo").Range.Text = " " + txtDiametroFondo.Text
        Documento.Bookmarks.Item("CATorno").Range.Text = " " + txtCATorno.Text
        Documento.Bookmarks.Item("APrimitivo").Range.Text = " " + txtAPrimitivo.Text
        Documento.Bookmarks.Item("AFondo").Range.Text = " " + txtAFondo.Text
        Documento.Bookmarks.Item("HAngulo").Range.Text = " " + txtHAngulo.Text
        Documento.Bookmarks.Item("helice").Range.Text = " " + txtHelice.Text
        Documento.Bookmarks.Item("CDiente").Range.Text = " " + txtCrestaDiente.Text
        Documento.Bookmarks.Item("ADiente").Range.Text = " " + txtAlturaDiente.Text
        Documento.Bookmarks.Item("fecha").Range.Text = " " + txtFecha.Text
        'Documento.Bookmarks.Item("fresadora").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("Pfresador").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("A1").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("A").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("plato").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("vueltas").Range.Text = " " + txtModuloOblicuo.Text

        'Documento.Bookmarks.Item("huecos").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("CGrueso").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("ESFin").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("DLlanta").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("ALaterales").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("ALlanta").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("RTransmision").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("tornillo").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("aparato").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("intermedios").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("N1").Range.Text = " " + txtModuloOblicuo.Text

        'Documento.Bookmarks.Item("N2").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("b").Range.Text = " " + txtModuloOblicuo.Text
        'Documento.Bookmarks.Item("c").Range.Text = " " + txtModuloOblicuo.Text
        Documento.Bookmarks.Item("hecho").Range.Text = " " + txtHechoPor.Text
        Documento.Bookmarks.Item("observaciones").Range.Text = " " + txtObservaciones.Text

        'Documento.Range().Font.Bold = False

        Documento.Save()
        MSWord.Visible = True
    End Sub

    Private Sub D4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class
