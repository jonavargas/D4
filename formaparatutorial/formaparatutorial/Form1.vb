Imports Microsoft.Office.Interop.Word 'control de office
Imports System.IO 'sistema de archivos
Imports Microsoft.Office.Interop

Public Class D4

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click
        Dim MSWord As New Word.Application
        Dim Documento As Word.Document

        MsgBox("El TDR se guardará en : C:\Users\Jonathan\Desktop\salida.doc")
        FileCopy("D:\documento.doc", "C:\Users\Jonathan\Desktop\salida.doc")
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

        Documento.Bookmarks.Item("fresadora").Range.Text = " " + txtFresadora.Text
        Documento.Bookmarks.Item("Pfresador").Range.Text = " " + txtPasoFresadora.Text
        Documento.Bookmarks.Item("A1").Range.Text = " " + txt40A1.Text
        Documento.Bookmarks.Item("A").Range.Text = " " + txtA.Text
        Documento.Bookmarks.Item("plato").Range.Text = " " + txtPlato.Text
        Documento.Bookmarks.Item("vueltas").Range.Text = " " + txtVueltas.Text

        Documento.Bookmarks.Item("huecos").Range.Text = " " + txtHueco.Text
        Documento.Bookmarks.Item("CGrueso").Range.Text = " " + txtCoronaGrueso.Text
        Documento.Bookmarks.Item("ESFin").Range.Text = " " + txtESFin.Text
        Documento.Bookmarks.Item("DLlanta").Range.Text = " " + txtDiametroLlanta.Text
        Documento.Bookmarks.Item("ALaterales").Range.Text = " " + txtAnguloLaterales.Text
        Documento.Bookmarks.Item("ALlanta").Range.Text = " " + txtAnchoLlanta.Text
        Documento.Bookmarks.Item("RTransmision").Range.Text = " " + txtRelacionTransmicion.Text
        Documento.Bookmarks.Item("tornillo").Range.Text = " " + txtTornilloA.Text
        Documento.Bookmarks.Item("aparato").Range.Text = " " + txtAparatoD.Text
        Documento.Bookmarks.Item("intermedios").Range.Text = " " + txtIntermedio1.Text
        Documento.Bookmarks.Item("N1").Range.Text = " " + txtModuloOblicuo.Text

        Documento.Bookmarks.Item("N2").Range.Text = " " + txtIntemedio2.Text
        Documento.Bookmarks.Item("b").Range.Text = " " + txtIntemedioB.Text
        Documento.Bookmarks.Item("c").Range.Text = " " + txtIntemedioC.Text
        Documento.Bookmarks.Item("hecho").Range.Text = " " + txtHechoPor.Text
        Documento.Bookmarks.Item("observaciones").Range.Text = " " + txtObservaciones.Text

        Documento.Save()
        MSWord.Visible = True
    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Limpiar_Cajas(Me)
        txtCliente.Focus()
    End Sub

    Private Sub Limpiar_Cajas(ByVal f As Form)
        txtCliente.Clear()
        txtModulo.Clear()
        txtPaso.Clear()
        txtModuloOblicuo.Clear()
        txtNumFresa.Clear()
        txtNumDientes.Clear()
        txtTipoEngranaje.Clear()
        txtDiametroExterior.Clear()
        txtDiametroPrimitivo.Clear()
        txtDiametroFondo.Clear()
        txtCATorno.Clear()
        txtAPrimitivo.Clear()
        txtAFondo.Clear()
        txtHAngulo.Clear()
        txtHelice.Clear()
        txtCrestaDiente.Clear()
        txtAlturaDiente.Clear()

        txtFresadora.Clear()
        txtPasoFresadora.Clear()
        txt40A1.Clear()
        txtA.Clear()
        txtPlato.Clear()
        txtVueltas.Clear()

        txtHueco.Clear()
        txtCoronaGrueso.Clear()
        txtESFin.Clear()
        txtDiametroLlanta.Clear()
        txtAnguloLaterales.Clear()
        txtAnchoLlanta.Clear()
        txtRelacionTransmicion.Clear()
        txtTornilloA.Clear()
        txtAparatoD.Clear()
        txtIntermedio1.Clear()
        txtModuloOblicuo.Clear()

        txtIntemedio2.Clear()
        txtIntemedioB.Clear()
        txtIntemedioC.Clear()
        txtHechoPor.Clear()
        txtObservaciones.Clear()
    End Sub
End Class
