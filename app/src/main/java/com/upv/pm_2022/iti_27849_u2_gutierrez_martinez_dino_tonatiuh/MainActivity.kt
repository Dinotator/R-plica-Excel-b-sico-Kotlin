package com.upv.pm_2022.iti_27849_u2_gutierrez_martinez_dino_tonatiuh


import android.Manifest
import android.content.Context
import android.content.pm.PackageManager
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.view.LayoutInflater
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import com.obsez.android.lib.filechooser.ChooserDialog
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


class MainActivity : AppCompatActivity(), View.OnClickListener {


    private val REQUEST_ID_READ_PERMISSION = 100
    private val REQUEST_ID_WRITE_PERMISSION = 200 //Estos son los permisos
    private var CX: Context? = null
    private val filePath: File =
        File(Environment.getExternalStorageDirectory().toString() + "/ArchivoGenerado.xlsx")

    var Ruta = ""
    var startingDir = Environment.getExternalStorageDirectory().toString()
    var datoCelda = ArrayList<String>()
    var nombreCelda = ArrayList<String>()
    var textViewArray = ArrayList<TextView>()

    lateinit var a1: TextView
    lateinit var b1: TextView
    lateinit var c1: TextView
    lateinit var d1: TextView
    lateinit var e1: TextView
    lateinit var f1: TextView
    lateinit var g1: TextView
    lateinit var h1: TextView
    lateinit var i1: TextView
    lateinit var j1: TextView

    lateinit var a2: TextView
    lateinit var b2: TextView
    lateinit var c2: TextView
    lateinit var d2: TextView
    lateinit var e2: TextView
    lateinit var f2: TextView
    lateinit var g2: TextView
    lateinit var h2: TextView
    lateinit var i2: TextView
    lateinit var j2: TextView

    lateinit var a3: TextView
    lateinit var b3: TextView
    lateinit var c3: TextView
    lateinit var d3: TextView
    lateinit var e3: TextView
    lateinit var f3: TextView
    lateinit var g3: TextView
    lateinit var h3: TextView
    lateinit var i3: TextView
    lateinit var j3: TextView

    lateinit var a4: TextView
    lateinit var b4: TextView
    lateinit var c4: TextView
    lateinit var d4: TextView
    lateinit var e4: TextView
    lateinit var f4: TextView
    lateinit var g4: TextView
    lateinit var h4: TextView
    lateinit var i4: TextView
    lateinit var j4: TextView

    lateinit var a5: TextView
    lateinit var b5: TextView
    lateinit var c5: TextView
    lateinit var d5: TextView
    lateinit var e5: TextView
    lateinit var f5: TextView
    lateinit var g5: TextView
    lateinit var h5: TextView
    lateinit var i5: TextView
    lateinit var j5: TextView

    lateinit var a6: TextView
    lateinit var b6: TextView
    lateinit var c6: TextView
    lateinit var d6: TextView
    lateinit var e6: TextView
    lateinit var f6: TextView
    lateinit var g6: TextView
    lateinit var h6: TextView
    lateinit var i6: TextView
    lateinit var j6: TextView

    lateinit var a7: TextView
    lateinit var b7: TextView
    lateinit var c7: TextView
    lateinit var d7: TextView
    lateinit var e7: TextView
    lateinit var f7: TextView
    lateinit var g7: TextView
    lateinit var h7: TextView
    lateinit var i7: TextView
    lateinit var j7: TextView

    lateinit var a8: TextView
    lateinit var b8: TextView
    lateinit var c8: TextView
    lateinit var d8: TextView
    lateinit var e8: TextView
    lateinit var f8: TextView
    lateinit var g8: TextView
    lateinit var h8: TextView
    lateinit var i8: TextView
    lateinit var j8: TextView

    lateinit var a9: TextView
    lateinit var b9: TextView
    lateinit var c9: TextView
    lateinit var d9: TextView
    lateinit var e9: TextView
    lateinit var f9: TextView
    lateinit var g9: TextView
    lateinit var h9: TextView
    lateinit var i9: TextView
    lateinit var j9: TextView

    lateinit var a10: TextView
    lateinit var b10: TextView
    lateinit var c10: TextView
    lateinit var d10: TextView
    lateinit var e10: TextView
    lateinit var f10: TextView
    lateinit var g10: TextView
    lateinit var h10: TextView
    lateinit var i10: TextView
    lateinit var j10: TextView

    lateinit var a11: TextView
    lateinit var b11: TextView
    lateinit var c11: TextView
    lateinit var d11: TextView
    lateinit var e11: TextView
    lateinit var f11: TextView
    lateinit var g11: TextView
    lateinit var h11: TextView
    lateinit var i11: TextView
    lateinit var j11: TextView

    lateinit var a12: TextView
    lateinit var b12: TextView
    lateinit var c12: TextView
    lateinit var d12: TextView
    lateinit var e12: TextView
    lateinit var f12: TextView
    lateinit var g12: TextView
    lateinit var h12: TextView
    lateinit var i12: TextView
    lateinit var j12: TextView

    lateinit var a13: TextView
    lateinit var b13: TextView
    lateinit var c13: TextView
    lateinit var d13: TextView
    lateinit var e13: TextView
    lateinit var f13: TextView
    lateinit var g13: TextView
    lateinit var h13: TextView
    lateinit var i13: TextView
    lateinit var j13: TextView

    lateinit var a14: TextView
    lateinit var b14: TextView
    lateinit var c14: TextView
    lateinit var d14: TextView
    lateinit var e14: TextView
    lateinit var f14: TextView
    lateinit var g14: TextView
    lateinit var h14: TextView
    lateinit var i14: TextView
    lateinit var j14: TextView

    lateinit var a15: TextView
    lateinit var b15: TextView
    lateinit var c15: TextView
    lateinit var d15: TextView
    lateinit var e15: TextView
    lateinit var f15: TextView
    lateinit var g15: TextView
    lateinit var h15: TextView
    lateinit var i15: TextView
    lateinit var j15: TextView

    lateinit var a16: TextView
    lateinit var b16: TextView
    lateinit var c16: TextView
    lateinit var d16: TextView
    lateinit var e16: TextView
    lateinit var f16: TextView
    lateinit var g16: TextView
    lateinit var h16: TextView
    lateinit var i16: TextView
    lateinit var j16: TextView

    lateinit var a17: TextView
    lateinit var b17: TextView
    lateinit var c17: TextView
    lateinit var d17: TextView
    lateinit var e17: TextView
    lateinit var f17: TextView
    lateinit var g17: TextView
    lateinit var h17: TextView
    lateinit var i17: TextView
    lateinit var j17: TextView

    lateinit var a18: TextView
    lateinit var b18: TextView
    lateinit var c18: TextView
    lateinit var d18: TextView
    lateinit var e18: TextView
    lateinit var f18: TextView
    lateinit var g18: TextView
    lateinit var h18: TextView
    lateinit var i18: TextView
    lateinit var j18: TextView

    lateinit var a19: TextView
    lateinit var b19: TextView
    lateinit var c19: TextView
    lateinit var d19: TextView
    lateinit var e19: TextView
    lateinit var f19: TextView
    lateinit var g19: TextView
    lateinit var h19: TextView
    lateinit var i19: TextView
    lateinit var j19: TextView

    lateinit var a20: TextView
    lateinit var b20: TextView
    lateinit var c20: TextView
    lateinit var d20: TextView
    lateinit var e20: TextView
    lateinit var f20: TextView
    lateinit var g20: TextView
    lateinit var h20: TextView
    lateinit var i20: TextView
    lateinit var j20: TextView


    lateinit var btnImportar: Button
    lateinit var btnGuardar: Button
    lateinit var btnLimpiar: Button


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        askPermissionOnly()
        CX = this;
        btnGuardar = findViewById(R.id.button_guardar)
        btnImportar = findViewById(R.id.button_importar)
        btnLimpiar = findViewById(R.id.button_limpiar)


        a1 = findViewById(R.id.a1);b1 = findViewById(R.id.b1); c1 = findViewById(R.id.c1); d1 =
            findViewById(R.id.d1); e1 = findViewById(R.id.e1); f1 = findViewById(R.id.f1); g1 =
            findViewById(R.id.g1); h1 = findViewById(R.id.h1); i1 = findViewById(R.id.i1); j1 =
            findViewById(R.id.j1)
        a2 = findViewById(R.id.a2);b2 = findViewById(R.id.b2); c2 = findViewById(R.id.c2); d2 =
            findViewById(R.id.d2); e2 = findViewById(R.id.e2); f2 = findViewById(R.id.f2); g2 =
            findViewById(R.id.g2); h2 = findViewById(R.id.h2); i2 = findViewById(R.id.i2); j2 =
            findViewById(R.id.j2)
        a3 = findViewById(R.id.a3);b3 = findViewById(R.id.b3); c3 = findViewById(R.id.c3); d3 =
            findViewById(R.id.d3); e3 = findViewById(R.id.e3); f3 = findViewById(R.id.f3); g3 =
            findViewById(R.id.g3); h3 = findViewById(R.id.h3); i3 = findViewById(R.id.i3); j3 =
            findViewById(R.id.j3)
        a4 = findViewById(R.id.a4);b4 = findViewById(R.id.b4); c4 = findViewById(R.id.c4); d4 =
            findViewById(R.id.d4); e4 = findViewById(R.id.e4); f4 = findViewById(R.id.f4); g4 =
            findViewById(R.id.g4); h4 = findViewById(R.id.h4); i4 = findViewById(R.id.i4); j4 =
            findViewById(R.id.j4)
        a5 = findViewById(R.id.a5);b5 = findViewById(R.id.b5); c5 = findViewById(R.id.c5); d5 =
            findViewById(R.id.d5); e5 = findViewById(R.id.e5); f5 = findViewById(R.id.f5); g5 =
            findViewById(R.id.g5); h5 = findViewById(R.id.h5); i5 = findViewById(R.id.i5); j5 =
            findViewById(R.id.j5)
        a6 = findViewById(R.id.a6);b6 = findViewById(R.id.b6); c6 = findViewById(R.id.c6); d6 =
            findViewById(R.id.d6); e6 = findViewById(R.id.e6); f6 = findViewById(R.id.f6); g6 =
            findViewById(R.id.g6); h6 = findViewById(R.id.h6); i6 = findViewById(R.id.i6); j6 =
            findViewById(R.id.j6)
        a7 = findViewById(R.id.a7);b7 = findViewById(R.id.b7); c7 = findViewById(R.id.c7); d7 =
            findViewById(R.id.d7); e7 = findViewById(R.id.e7); f7 = findViewById(R.id.f7); g7 =
            findViewById(R.id.g7); h7 = findViewById(R.id.h7); i7 = findViewById(R.id.i7); j7 =
            findViewById(R.id.j7)
        a8 = findViewById(R.id.a8);b8 = findViewById(R.id.b8); c8 = findViewById(R.id.c8); d8 =
            findViewById(R.id.d8); e8 = findViewById(R.id.e8); f8 = findViewById(R.id.f8); g8 =
            findViewById(R.id.g8); h8 = findViewById(R.id.h8); i8 = findViewById(R.id.i8); j8 =
            findViewById(R.id.j8)
        a9 = findViewById(R.id.a9);b9 = findViewById(R.id.b9); c9 = findViewById(R.id.c9); d9 =
            findViewById(R.id.d9); e9 = findViewById(R.id.e9); f9 = findViewById(R.id.f9); g9 =
            findViewById(R.id.g9); h9 = findViewById(R.id.h9); i9 = findViewById(R.id.i9); j9 =
            findViewById(R.id.j9)
        a10 = findViewById(R.id.a10);b10 = findViewById(R.id.b10); c10 =
            findViewById(R.id.c10); d10 = findViewById(R.id.d10); e10 =
            findViewById(R.id.e10); f10 = findViewById(R.id.f10); g10 =
            findViewById(R.id.g10); h10 = findViewById(R.id.h10); i10 =
            findViewById(R.id.i10); j10 = findViewById(R.id.j10)

        a11 = findViewById(R.id.a11);b11 = findViewById(R.id.b11); c11 =
            findViewById(R.id.c11); d11 = findViewById(R.id.d11); e11 =
            findViewById(R.id.e11); f11 = findViewById(R.id.f11); g11 =
            findViewById(R.id.g11); h11 = findViewById(R.id.h11); i11 =
            findViewById(R.id.i11); j11 = findViewById(R.id.j11)
        a12 = findViewById(R.id.a12);b12 = findViewById(R.id.b12); c12 =
            findViewById(R.id.c12); d12 = findViewById(R.id.d12); e12 =
            findViewById(R.id.e12); f12 = findViewById(R.id.f12); g12 =
            findViewById(R.id.g12); h12 = findViewById(R.id.h12); i12 =
            findViewById(R.id.i12); j12 = findViewById(R.id.j12)
        a13 = findViewById(R.id.a13);b13 = findViewById(R.id.b13); c13 =
            findViewById(R.id.c13); d13 = findViewById(R.id.d13); e13 =
            findViewById(R.id.e13); f13 = findViewById(R.id.f13); g13 =
            findViewById(R.id.g13); h13 = findViewById(R.id.h13); i13 =
            findViewById(R.id.i13); j13 = findViewById(R.id.j13)
        a14 = findViewById(R.id.a14);b14 = findViewById(R.id.b14); c14 =
            findViewById(R.id.c14); d14 = findViewById(R.id.d14); e14 =
            findViewById(R.id.e14); f14 = findViewById(R.id.f14); g14 =
            findViewById(R.id.g14); h14 = findViewById(R.id.h14); i14 =
            findViewById(R.id.i14); j14 = findViewById(R.id.j14)
        a15 = findViewById(R.id.a15);b15 = findViewById(R.id.b15); c15 =
            findViewById(R.id.c15); d15 = findViewById(R.id.d15); e15 =
            findViewById(R.id.e15); f15 = findViewById(R.id.f15); g15 =
            findViewById(R.id.g15); h15 = findViewById(R.id.h15); i15 =
            findViewById(R.id.i15); j15 = findViewById(R.id.j15)
        a16 = findViewById(R.id.a16);b16 = findViewById(R.id.b16); c16 =
            findViewById(R.id.c16); d16 = findViewById(R.id.d16); e16 =
            findViewById(R.id.e16); f16 = findViewById(R.id.f16); g16 =
            findViewById(R.id.g16); h16 = findViewById(R.id.h16); i16 =
            findViewById(R.id.i16); j16 = findViewById(R.id.j16)
        a17 = findViewById(R.id.a17);b17 = findViewById(R.id.b17); c17 =
            findViewById(R.id.c17); d17 = findViewById(R.id.d17); e17 =
            findViewById(R.id.e17); f17 = findViewById(R.id.f17); g17 =
            findViewById(R.id.g17); h17 = findViewById(R.id.h17); i17 =
            findViewById(R.id.i17); j17 = findViewById(R.id.j17)
        a18 = findViewById(R.id.a18);b18 = findViewById(R.id.b18); c18 =
            findViewById(R.id.c18); d18 = findViewById(R.id.d18); e18 =
            findViewById(R.id.e18); f18 = findViewById(R.id.f18); g18 =
            findViewById(R.id.g18); h18 = findViewById(R.id.h18); i18 =
            findViewById(R.id.i18); j18 = findViewById(R.id.j18)
        a19 = findViewById(R.id.a19);b19 = findViewById(R.id.b19); c19 =
            findViewById(R.id.c19); d19 = findViewById(R.id.d19); e19 =
            findViewById(R.id.e19); f19 = findViewById(R.id.f19); g19 =
            findViewById(R.id.g19); h19 = findViewById(R.id.h19); i19 =
            findViewById(R.id.i19); j19 = findViewById(R.id.j19)
        a20 = findViewById(R.id.a20);b20 = findViewById(R.id.b20); c20 =
            findViewById(R.id.c20); d20 = findViewById(R.id.d20); e20 =
            findViewById(R.id.e20); f20 = findViewById(R.id.f20); g20 =
            findViewById(R.id.g20); h20 = findViewById(R.id.h20); i20 =
            findViewById(R.id.i20); j20 = findViewById(R.id.j20)


        //Declaramos un arreglo el cual almacenará el contenido en tipo cadena de todos los TextView
        datoCelda = arrayListOf(
            a1.text.toString(),
            b1.text.toString(),
            c1.text.toString(),
            d1.text.toString(),
            e1.text.toString(),
            f1.text.toString(),
            g1.text.toString(),
            h1.text.toString(),
            i1.text.toString(),
            j1.text.toString(),

            a2.text.toString(),
            b2.text.toString(),
            c2.text.toString(),
            d2.text.toString(),
            e2.text.toString(),
            f2.text.toString(),
            g2.text.toString(),
            h2.text.toString(),
            i2.text.toString(),
            j2.text.toString(),

            a3.text.toString(),
            b3.text.toString(),
            c3.text.toString(),
            d3.text.toString(),
            e3.text.toString(),
            f3.text.toString(),
            g3.text.toString(),
            h3.text.toString(),
            i3.text.toString(),
            j3.text.toString(),

            a4.text.toString(),
            b4.text.toString(),
            c4.text.toString(),
            d4.text.toString(),
            e4.text.toString(),
            f4.text.toString(),
            g4.text.toString(),
            h4.text.toString(),
            i4.text.toString(),
            j4.text.toString(),

            a5.text.toString(),
            b5.text.toString(),
            c5.text.toString(),
            d5.text.toString(),
            e5.text.toString(),
            f5.text.toString(),
            g5.text.toString(),
            h5.text.toString(),
            i5.text.toString(),
            j5.text.toString(),

            a6.text.toString(),
            b6.text.toString(),
            c6.text.toString(),
            d6.text.toString(),
            e6.text.toString(),
            f6.text.toString(),
            g6.text.toString(),
            h6.text.toString(),
            i6.text.toString(),
            j6.text.toString(),

            a7.text.toString(),
            b7.text.toString(),
            c7.text.toString(),
            d7.text.toString(),
            e7.text.toString(),
            f7.text.toString(),
            g7.text.toString(),
            h7.text.toString(),
            i7.text.toString(),
            j7.text.toString(),

            a8.text.toString(),
            b8.text.toString(),
            c8.text.toString(),
            d8.text.toString(),
            e8.text.toString(),
            f8.text.toString(),
            g8.text.toString(),
            h8.text.toString(),
            i8.text.toString(),
            j8.text.toString(),

            a9.text.toString(),
            b9.text.toString(),
            c9.text.toString(),
            d9.text.toString(),
            e9.text.toString(),
            f9.text.toString(),
            g9.text.toString(),
            h9.text.toString(),
            i9.text.toString(),
            j9.text.toString(),

            a10.text.toString(),
            b10.text.toString(),
            c10.text.toString(),
            d10.text.toString(),
            e10.text.toString(),
            f10.text.toString(),
            g10.text.toString(),
            h10.text.toString(),
            i10.text.toString(),
            j10.text.toString(),

            a11.text.toString(),
            b11.text.toString(),
            c11.text.toString(),
            d11.text.toString(),
            e11.text.toString(),
            f11.text.toString(),
            g11.text.toString(),
            h11.text.toString(),
            i11.text.toString(),
            j11.text.toString(),

            a12.text.toString(),
            b12.text.toString(),
            c12.text.toString(),
            d12.text.toString(),
            e12.text.toString(),
            f12.text.toString(),
            g12.text.toString(),
            h12.text.toString(),
            i12.text.toString(),
            j12.text.toString(),

            a13.text.toString(),
            b13.text.toString(),
            c13.text.toString(),
            d13.text.toString(),
            e13.text.toString(),
            f13.text.toString(),
            g13.text.toString(),
            h13.text.toString(),
            i13.text.toString(),
            j13.text.toString(),

            a14.text.toString(),
            b14.text.toString(),
            c14.text.toString(),
            d14.text.toString(),
            e14.text.toString(),
            f14.text.toString(),
            g14.text.toString(),
            h14.text.toString(),
            i14.text.toString(),
            j14.text.toString(),

            a15.text.toString(),
            b15.text.toString(),
            c15.text.toString(),
            d15.text.toString(),
            e15.text.toString(),
            f15.text.toString(),
            g15.text.toString(),
            h15.text.toString(),
            i15.text.toString(),
            j15.text.toString(),

            a16.text.toString(),
            b16.text.toString(),
            c16.text.toString(),
            d16.text.toString(),
            e16.text.toString(),
            f16.text.toString(),
            g16.text.toString(),
            h16.text.toString(),
            i16.text.toString(),
            j16.text.toString(),

            a17.text.toString(),
            b17.text.toString(),
            c17.text.toString(),
            d17.text.toString(),
            e17.text.toString(),
            f17.text.toString(),
            g17.text.toString(),
            h17.text.toString(),
            i17.text.toString(),
            j17.text.toString(),

            a18.text.toString(),
            b18.text.toString(),
            c18.text.toString(),
            d18.text.toString(),
            e18.text.toString(),
            f18.text.toString(),
            g18.text.toString(),
            h18.text.toString(),
            i18.text.toString(),
            j18.text.toString(),

            a19.text.toString(),
            b19.text.toString(),
            c19.text.toString(),
            d19.text.toString(),
            e19.text.toString(),
            f19.text.toString(),
            g19.text.toString(),
            h19.text.toString(),
            i19.text.toString(),
            j19.text.toString(),

            a20.text.toString(),
            b20.text.toString(),
            c20.text.toString(),
            d20.text.toString(),
            e20.text.toString(),
            f20.text.toString(),
            g20.text.toString(),
            h20.text.toString(),
            i20.text.toString(),
            j20.text.toString()

        )

        //Declaramos un arreglo el cual almacenará los objetos TextView
        textViewArray = arrayListOf(
            a1, b1, c1, d1, e1, f1, g1, h1, i1, j1,
            a2, b2, c2, d2, e2, f2, g2, h2, i2, j2,
            a3, b3, c3, d3, e3, f3, g3, h3, i3, j3,
            a4, b4, c4, d4, e4, f4, g4, h4, i4, j4,
            a5, b5, c5, d5, e5, f5, g5, h5, i5, j5,
            a6, b6, c6, d6, e6, f6, g6, h6, i6, j6,
            a7, b7, c7, d7, e7, f7, g7, h7, i7, j7,
            a8, b8, c8, d8, e8, f8, g8, h8, i8, j8,
            a9, b9, c9, d9, e9, f9, g9, h9, i9, j9,
            a10, b10, c10, d10, e10, f10, g10, h10, i10, j10,

            a11, b11, c11, d11, e11, f11, g11, h11, i11, j11,
            a12, b12, c12, d12, e12, f12, g12, h12, i12, j12,
            a13, b13, c13, d13, e13, f13, g13, h13, i13, j13,
            a14, b14, c14, d14, e14, f14, g14, h14, i14, j14,
            a15, b15, c15, d15, e15, f15, g15, h15, i15, j15,
            a16, b16, c16, d16, e16, f16, g16, h16, i16, j16,
            a17, b17, c17, d17, e17, f17, g17, h17, i17, j17,
            a18, b18, c18, d18, e18, f18, g18, h18, i18, j18,
            a19, b19, c19, d19, e19, f19, g19, h19, i19, j19,
            a20, b20, c20, d20, e20, f20, g20, h20, i20, j20

        )

        nombreCelda = arrayListOf(
            "A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1",
            "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2",
            "A3", "B3", "C3", "D3", "E3", "F3", "G3", "H3", "I3", "J3",
            "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4",
            "A5", "B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5",
            "A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6",
            "A7", "B7", "C7", "D7", "E7", "F7", "G7", "H7", "I7", "J7",
            "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "J8",
            "A9", "B9", "C9", "D9", "E9", "F9", "G9", "H9", "I9", "J9",
            "A10", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10",

            "A11", "B11", "C11", "D11", "E11", "F11", "G11", "H11", "I11", "J11",
            "A12", "B12", "C12", "D12", "E12", "F12", "G12", "H12", "I12", "J12",
            "A13", "B13", "C13", "D13", "E13", "F13", "G13", "H13", "I13", "J13",
            "A14", "B14", "C14", "D14", "E14", "F14", "G14", "H14", "I14", "J14",
            "A15", "B15", "C15", "D15", "E15", "F15", "G15", "H15", "I15", "J15",
            "A16", "B16", "C16", "D16", "E16", "F16", "G16", "H16", "I16", "J16",
            "A17", "B17", "C17", "D17", "E17", "F17", "G17", "H17", "I17", "J17",
            "A18", "B18", "C18", "D18", "E18", "F18", "G18", "H18", "I18", "J18",
            "A19", "B19", "C19", "D19", "E19", "F19", "G19", "H19", "I19", "J19",
            "A20", "B20", "C20", "D20", "E20", "F20", "G20", "H20", "I20", "J20",

            )


        //Se asignan a todos los TextView Los setOnClickListener, con la finalidad de poner meterlos a un "Switch"
        //El cual dependiendo de cual se escoja, va a ser con el que se esté manipulando
        a1.setOnClickListener(this);b1.setOnClickListener(this);c1.setOnClickListener(this);d1.setOnClickListener(
            this
        );e1.setOnClickListener(this);f1.setOnClickListener(this);g1.setOnClickListener(this);h1.setOnClickListener(
            this
        );i1.setOnClickListener(this);j1.setOnClickListener(this);
        a2.setOnClickListener(this);b2.setOnClickListener(this);c2.setOnClickListener(this);d2.setOnClickListener(
            this
        );e2.setOnClickListener(this);f2.setOnClickListener(this);g2.setOnClickListener(this);h2.setOnClickListener(
            this
        );i2.setOnClickListener(this);j2.setOnClickListener(this);
        a3.setOnClickListener(this);b3.setOnClickListener(this);c3.setOnClickListener(this);d3.setOnClickListener(
            this
        );e3.setOnClickListener(this);f3.setOnClickListener(this);g3.setOnClickListener(this);h3.setOnClickListener(
            this
        );i3.setOnClickListener(this);j3.setOnClickListener(this);
        a4.setOnClickListener(this);b4.setOnClickListener(this);c4.setOnClickListener(this);d4.setOnClickListener(
            this
        );e4.setOnClickListener(this);f4.setOnClickListener(this);g4.setOnClickListener(this);h4.setOnClickListener(
            this
        );i4.setOnClickListener(this);j4.setOnClickListener(this);
        a5.setOnClickListener(this);b5.setOnClickListener(this);c5.setOnClickListener(this);d5.setOnClickListener(
            this
        );e5.setOnClickListener(this);f5.setOnClickListener(this);g5.setOnClickListener(this);h5.setOnClickListener(
            this
        );i5.setOnClickListener(this);j5.setOnClickListener(this);
        a6.setOnClickListener(this);b6.setOnClickListener(this);c6.setOnClickListener(this);d6.setOnClickListener(
            this
        );e6.setOnClickListener(this);f6.setOnClickListener(this);g6.setOnClickListener(this);h6.setOnClickListener(
            this
        );i6.setOnClickListener(this);j6.setOnClickListener(this);
        a7.setOnClickListener(this);b7.setOnClickListener(this);c7.setOnClickListener(this);d7.setOnClickListener(
            this
        );e7.setOnClickListener(this);f7.setOnClickListener(this);g7.setOnClickListener(this);h7.setOnClickListener(
            this
        );i7.setOnClickListener(this);j7.setOnClickListener(this);
        a8.setOnClickListener(this);b8.setOnClickListener(this);c8.setOnClickListener(this);d8.setOnClickListener(
            this
        );e8.setOnClickListener(this);f8.setOnClickListener(this);g8.setOnClickListener(this);h8.setOnClickListener(
            this
        );i8.setOnClickListener(this);j8.setOnClickListener(this);
        a9.setOnClickListener(this);b9.setOnClickListener(this);c9.setOnClickListener(this);d9.setOnClickListener(
            this
        );e9.setOnClickListener(this);f9.setOnClickListener(this);g9.setOnClickListener(this);h9.setOnClickListener(
            this
        );i9.setOnClickListener(this);j9.setOnClickListener(this);
        a10.setOnClickListener(this);b10.setOnClickListener(this);c10.setOnClickListener(this);d10.setOnClickListener(
            this
        );e10.setOnClickListener(this);f10.setOnClickListener(this);g10.setOnClickListener(this);h10.setOnClickListener(
            this
        );i10.setOnClickListener(this);j10.setOnClickListener(this);

        a11.setOnClickListener(this);b11.setOnClickListener(this);c11.setOnClickListener(this);d11.setOnClickListener(
            this
        );e11.setOnClickListener(this);f11.setOnClickListener(this);g11.setOnClickListener(this);h11.setOnClickListener(
            this
        );i11.setOnClickListener(this);j11.setOnClickListener(this);
        a12.setOnClickListener(this);b12.setOnClickListener(this);c12.setOnClickListener(this);d12.setOnClickListener(
            this
        );e12.setOnClickListener(this);f12.setOnClickListener(this);g12.setOnClickListener(this);h12.setOnClickListener(
            this
        );i12.setOnClickListener(this);j12.setOnClickListener(this);
        a13.setOnClickListener(this);b13.setOnClickListener(this);c13.setOnClickListener(this);d13.setOnClickListener(
            this
        );e13.setOnClickListener(this);f13.setOnClickListener(this);g13.setOnClickListener(this);h13.setOnClickListener(
            this
        );i13.setOnClickListener(this);j13.setOnClickListener(this);
        a14.setOnClickListener(this);b14.setOnClickListener(this);c14.setOnClickListener(this);d14.setOnClickListener(
            this
        );e14.setOnClickListener(this);f14.setOnClickListener(this);g14.setOnClickListener(this);h14.setOnClickListener(
            this
        );i14.setOnClickListener(this);j14.setOnClickListener(this);
        a15.setOnClickListener(this);b15.setOnClickListener(this);c15.setOnClickListener(this);d15.setOnClickListener(
            this
        );e15.setOnClickListener(this);f15.setOnClickListener(this);g15.setOnClickListener(this);h15.setOnClickListener(
            this
        );i15.setOnClickListener(this);j15.setOnClickListener(this);
        a16.setOnClickListener(this);b16.setOnClickListener(this);c16.setOnClickListener(this);d16.setOnClickListener(
            this
        );e16.setOnClickListener(this);f16.setOnClickListener(this);g16.setOnClickListener(this);h16.setOnClickListener(
            this
        );i16.setOnClickListener(this);j16.setOnClickListener(this);
        a17.setOnClickListener(this);b17.setOnClickListener(this);c17.setOnClickListener(this);d17.setOnClickListener(
            this
        );e17.setOnClickListener(this);f17.setOnClickListener(this);g17.setOnClickListener(this);h17.setOnClickListener(
            this
        );i17.setOnClickListener(this);j17.setOnClickListener(this);
        a18.setOnClickListener(this);b18.setOnClickListener(this);c18.setOnClickListener(this);d18.setOnClickListener(
            this
        );e18.setOnClickListener(this);f18.setOnClickListener(this);g18.setOnClickListener(this);h18.setOnClickListener(
            this
        );i18.setOnClickListener(this);j18.setOnClickListener(this);
        a19.setOnClickListener(this);b19.setOnClickListener(this);c19.setOnClickListener(this);d19.setOnClickListener(
            this
        );e19.setOnClickListener(this);f19.setOnClickListener(this);g19.setOnClickListener(this);h19.setOnClickListener(
            this
        );i19.setOnClickListener(this);j19.setOnClickListener(this);
        a20.setOnClickListener(this);b20.setOnClickListener(this);c20.setOnClickListener(this);d20.setOnClickListener(
            this
        );e20.setOnClickListener(this);f20.setOnClickListener(this);g20.setOnClickListener(this);h20.setOnClickListener(
            this
        );i20.setOnClickListener(this);j20.setOnClickListener(this);

        btnLimpiar.setOnClickListener {
            for (i in (0 until 199)) {
                textViewArray[i].text = ""
                datoCelda[i] = ""
            }
        }


        //La función de este botón es guardar el archivo xlsx
        btnGuardar.setOnClickListener {
            try {
                //Creamos el archivo Excel
                var xlWb = XSSFWorkbook()

                //Creamos una hoja con el nombre Archivo Generado
                var xlWs = xlWb.createSheet("Archivo Generado")

                //Declaramos la fila y la celda
                var row = xlWs.createRow(0)
                var cell = row.createCell(0)

                //Declaramos un auxiliar para poder recorrer el ArrayList sin perder el conteo
                var aux = 0

                //El primer for es para crear las 5 filas
                for (i in (0 until 20)) {
                    row = xlWs.createRow(i)
                    //El segundo for es para ir creando las celdas
                    for (j in (0 until 10)) {
                        cell = row.createCell(j)
                        //Aquí establecemos el valor de la celda usando el ArrayList que guarda todos los valores
                        cell.setCellValue(datoCelda[aux])
                        //Por cada iteración se irá agregando 1 en el auxiliar, y éste no se reiniciará el conteo a 0 como el j
                        aux++
                    }
                }

                //Asignamos un Try Catch, el cual valida si el archivo existe, de no ser así crea uno
                try {
                    if (!filePath.exists()) {
                        filePath.createNewFile()
                    }
                    //Escribimos el contenido que acabamos de crear para luego cerrarlo
                    val fileOutputStream = FileOutputStream(filePath)
                    xlWb.write(fileOutputStream)

                    if (fileOutputStream != null) {
                        fileOutputStream.flush()
                        fileOutputStream.close()
                    }
                } catch (e: Exception) {

                }

                Toast.makeText(this@MainActivity, "Archivo guardado", Toast.LENGTH_SHORT).show()
            } catch (e: Exception) {
                Toast.makeText(this@MainActivity, "Error al guardar", Toast.LENGTH_SHORT).show()

            }


        }

        btnImportar.setOnClickListener {
            ChooserDialog(this@MainActivity)
                .withFilter(false, false, "xlsx", "XLSX")
                .withStartFile(startingDir)
                .withResources(R.string.app_name, R.string.yes_button, R.string.no_button)
                .withChosenListener { path, _ ->
                    Ruta = path

                    leerExcel()
                }
                .build()
                .show()
        }
    }

    private fun leerExcel() = try {
        val input = FileInputStream(Ruta)
        val xlWb = WorkbookFactory.create(input)
        val xlWs = xlWb.getSheetAt(0)
        var aux = 0

        for (i in (0 until 20)) {
            var row = xlWs.getRow(i)
            for (j in (0 until 10)) {
                var cell = row.getCell(j)
                if (cell == null || cell.cellType == Cell.CELL_TYPE_BLANK) {
                    textViewArray[aux].text = ""
                    datoCelda[aux] = ""
                } else {

                    textViewArray[aux].text = funcion(cell.toString())
                    datoCelda[aux] = funcion(cell.toString())
                }
                aux++
            }
        }
        Toast.makeText(this@MainActivity, "Archivo leído exitosamente", Toast.LENGTH_SHORT)
            .show()
    } catch (e: Exception) {
    }

    //Esta función lo que hace es que al momento de dar clic a cualquier celda (TextView) éste hará llamado a dos funciones
    //ingresarValor() <- para editar los TextView
    //Ahorita hago el otro xd
    override fun onClick(view: View) {
        when (view.id) {
            R.id.a1 -> ingresarValor("A", 0, 1)
            R.id.b1 -> ingresarValor("B", 1, 1)
            R.id.c1 -> ingresarValor("C", 2, 1)
            R.id.d1 -> ingresarValor("D", 3, 1)
            R.id.e1 -> ingresarValor("E", 4, 1)
            R.id.f1 -> ingresarValor("F", 5, 1)
            R.id.g1 -> ingresarValor("G", 6, 1)
            R.id.h1 -> ingresarValor("H", 7, 1)
            R.id.i1 -> ingresarValor("I", 8, 1)
            R.id.j1 -> ingresarValor("J", 9, 1)

            R.id.a2 -> ingresarValor("A", 10, 2)
            R.id.b2 -> ingresarValor("B", 11, 2)
            R.id.c2 -> ingresarValor("C", 12, 2)
            R.id.d2 -> ingresarValor("D", 13, 2)
            R.id.e2 -> ingresarValor("E", 14, 2)
            R.id.f2 -> ingresarValor("F", 15, 2)
            R.id.g2 -> ingresarValor("G", 16, 2)
            R.id.h2 -> ingresarValor("H", 17, 2)
            R.id.i2 -> ingresarValor("I", 18, 2)
            R.id.j2 -> ingresarValor("J", 19, 2)

            R.id.a3 -> ingresarValor("A", 20, 3)
            R.id.b3 -> ingresarValor("B", 21, 3)
            R.id.c3 -> ingresarValor("C", 22, 3)
            R.id.d3 -> ingresarValor("D", 23, 3)
            R.id.e3 -> ingresarValor("E", 24, 3)
            R.id.f3 -> ingresarValor("F", 25, 3)
            R.id.g3 -> ingresarValor("G", 26, 3)
            R.id.h3 -> ingresarValor("H", 27, 3)
            R.id.i3 -> ingresarValor("I", 28, 3)
            R.id.j3 -> ingresarValor("J", 29, 3)

            R.id.a4 -> ingresarValor("A", 30, 4)
            R.id.b4 -> ingresarValor("B", 31, 4)
            R.id.c4 -> ingresarValor("C", 32, 4)
            R.id.d4 -> ingresarValor("D", 33, 4)
            R.id.e4 -> ingresarValor("E", 34, 4)
            R.id.f4 -> ingresarValor("F", 35, 4)
            R.id.g4 -> ingresarValor("G", 36, 4)
            R.id.h4 -> ingresarValor("H", 37, 4)
            R.id.i4 -> ingresarValor("I", 38, 4)
            R.id.j4 -> ingresarValor("J", 39, 4)

            R.id.a5 -> ingresarValor("A", 40, 5)
            R.id.b5 -> ingresarValor("B", 41, 5)
            R.id.c5 -> ingresarValor("C", 42, 5)
            R.id.d5 -> ingresarValor("D", 43, 5)
            R.id.e5 -> ingresarValor("E", 44, 5)
            R.id.f5 -> ingresarValor("F", 45, 5)
            R.id.g5 -> ingresarValor("G", 46, 5)
            R.id.h5 -> ingresarValor("H", 47, 5)
            R.id.i5 -> ingresarValor("I", 48, 5)
            R.id.j5 -> ingresarValor("J", 49, 5)

            R.id.a6 -> ingresarValor("A", 50, 6)
            R.id.b6 -> ingresarValor("B", 51, 6)
            R.id.c6 -> ingresarValor("C", 52, 6)
            R.id.d6 -> ingresarValor("D", 53, 6)
            R.id.e6 -> ingresarValor("E", 54, 6)
            R.id.f6 -> ingresarValor("F", 55, 6)
            R.id.g6 -> ingresarValor("G", 56, 6)
            R.id.h6 -> ingresarValor("H", 57, 6)
            R.id.i6 -> ingresarValor("I", 58, 6)
            R.id.j6 -> ingresarValor("J", 59, 6)

            R.id.a7 -> ingresarValor("A", 60, 7)
            R.id.b7 -> ingresarValor("B", 61, 7)
            R.id.c7 -> ingresarValor("C", 62, 7)
            R.id.d7 -> ingresarValor("D", 63, 7)
            R.id.e7 -> ingresarValor("E", 64, 7)
            R.id.f7 -> ingresarValor("F", 65, 7)
            R.id.g7 -> ingresarValor("G", 66, 7)
            R.id.h7 -> ingresarValor("H", 67, 7)
            R.id.i7 -> ingresarValor("I", 68, 7)
            R.id.j7 -> ingresarValor("J", 69, 7)

            R.id.a8 -> ingresarValor("A", 70, 8)
            R.id.b8 -> ingresarValor("B", 71, 8)
            R.id.c8 -> ingresarValor("C", 72, 8)
            R.id.d8 -> ingresarValor("D", 73, 8)
            R.id.e8 -> ingresarValor("E", 74, 8)
            R.id.f8 -> ingresarValor("F", 75, 8)
            R.id.g8 -> ingresarValor("G", 76, 8)
            R.id.h8 -> ingresarValor("H", 77, 8)
            R.id.i8 -> ingresarValor("I", 78, 8)
            R.id.j8 -> ingresarValor("J", 79, 8)

            R.id.a9 -> ingresarValor("A", 80, 9)
            R.id.b9 -> ingresarValor("B", 81, 9)
            R.id.c9 -> ingresarValor("C", 82, 9)
            R.id.d9 -> ingresarValor("D", 83, 9)
            R.id.e9 -> ingresarValor("E", 84, 9)
            R.id.f9 -> ingresarValor("F", 85, 9)
            R.id.g9 -> ingresarValor("G", 86, 9)
            R.id.h9 -> ingresarValor("H", 87, 9)
            R.id.i9 -> ingresarValor("I", 88, 9)
            R.id.j9 -> ingresarValor("J", 89, 9)

            R.id.a10 -> ingresarValor("A", 90, 10)
            R.id.b10 -> ingresarValor("B", 91, 10)
            R.id.c10 -> ingresarValor("C", 92, 10)
            R.id.d10 -> ingresarValor("D", 93, 10)
            R.id.e10 -> ingresarValor("E", 94, 10)
            R.id.f10 -> ingresarValor("F", 95, 10)
            R.id.g10 -> ingresarValor("G", 96, 10)
            R.id.h10 -> ingresarValor("H", 97, 10)
            R.id.i10 -> ingresarValor("I", 98, 10)
            R.id.j10 -> ingresarValor("J", 99, 10)

            R.id.a11 -> ingresarValor("A", 100, 11)
            R.id.b11 -> ingresarValor("B", 101, 11)
            R.id.c11 -> ingresarValor("C", 102, 11)
            R.id.d11 -> ingresarValor("D", 103, 11)
            R.id.e11 -> ingresarValor("E", 104, 11)
            R.id.f11 -> ingresarValor("F", 105, 11)
            R.id.g11 -> ingresarValor("G", 106, 11)
            R.id.h11 -> ingresarValor("H", 107, 11)
            R.id.i11 -> ingresarValor("I", 108, 11)
            R.id.j11 -> ingresarValor("J", 109, 11)

            R.id.a12 -> ingresarValor("A", 110, 12)
            R.id.b12 -> ingresarValor("B", 111, 12)
            R.id.c12 -> ingresarValor("C", 112, 12)
            R.id.d12 -> ingresarValor("D", 113, 12)
            R.id.e12 -> ingresarValor("E", 114, 12)
            R.id.f12 -> ingresarValor("F", 115, 12)
            R.id.g12 -> ingresarValor("G", 116, 12)
            R.id.h12 -> ingresarValor("H", 117, 12)
            R.id.i12 -> ingresarValor("I", 118, 12)
            R.id.j12 -> ingresarValor("J", 119, 12)

            R.id.a13 -> ingresarValor("A", 120, 13)
            R.id.b13 -> ingresarValor("B", 121, 13)
            R.id.c13 -> ingresarValor("C", 122, 13)
            R.id.d13 -> ingresarValor("D", 123, 13)
            R.id.e13 -> ingresarValor("E", 124, 13)
            R.id.f13 -> ingresarValor("F", 125, 13)
            R.id.g13 -> ingresarValor("G", 126, 13)
            R.id.h13 -> ingresarValor("H", 127, 13)
            R.id.i13 -> ingresarValor("I", 128, 13)
            R.id.j13 -> ingresarValor("J", 129, 13)

            R.id.a14 -> ingresarValor("A", 130, 14)
            R.id.b14 -> ingresarValor("B", 131, 14)
            R.id.c14 -> ingresarValor("C", 132, 14)
            R.id.d14 -> ingresarValor("D", 133, 14)
            R.id.e14 -> ingresarValor("E", 134, 14)
            R.id.f14 -> ingresarValor("F", 135, 14)
            R.id.g14 -> ingresarValor("G", 136, 14)
            R.id.h14 -> ingresarValor("H", 137, 14)
            R.id.i14 -> ingresarValor("I", 138, 14)
            R.id.j14 -> ingresarValor("J", 139, 14)

            R.id.a15 -> ingresarValor("A", 140, 15)
            R.id.b15 -> ingresarValor("B", 141, 15)
            R.id.c15 -> ingresarValor("C", 142, 15)
            R.id.d15 -> ingresarValor("D", 143, 15)
            R.id.e15 -> ingresarValor("E", 144, 15)
            R.id.f15 -> ingresarValor("F", 145, 15)
            R.id.g15 -> ingresarValor("G", 146, 15)
            R.id.h15 -> ingresarValor("H", 147, 15)
            R.id.i15 -> ingresarValor("I", 148, 15)
            R.id.j15 -> ingresarValor("J", 149, 15)

            R.id.a16 -> ingresarValor("A", 150, 16)
            R.id.b16 -> ingresarValor("B", 151, 16)
            R.id.c16 -> ingresarValor("C", 152, 16)
            R.id.d16 -> ingresarValor("D", 153, 16)
            R.id.e16 -> ingresarValor("E", 154, 16)
            R.id.f16 -> ingresarValor("F", 155, 16)
            R.id.g16 -> ingresarValor("G", 156, 16)
            R.id.h16 -> ingresarValor("H", 157, 16)
            R.id.i16 -> ingresarValor("I", 158, 16)
            R.id.j16 -> ingresarValor("J", 159, 16)

            R.id.a17 -> ingresarValor("A", 160, 17)
            R.id.b17 -> ingresarValor("B", 161, 17)
            R.id.c17 -> ingresarValor("C", 162, 17)
            R.id.d17 -> ingresarValor("D", 163, 17)
            R.id.e17 -> ingresarValor("E", 164, 17)
            R.id.f17 -> ingresarValor("F", 165, 17)
            R.id.g17 -> ingresarValor("G", 166, 17)
            R.id.h17 -> ingresarValor("H", 167, 17)
            R.id.i17 -> ingresarValor("I", 168, 17)
            R.id.j17 -> ingresarValor("J", 169, 17)

            R.id.a18 -> ingresarValor("A", 170, 18)
            R.id.b18 -> ingresarValor("B", 171, 18)
            R.id.c18 -> ingresarValor("C", 172, 18)
            R.id.d18 -> ingresarValor("D", 173, 18)
            R.id.e18 -> ingresarValor("E", 174, 18)
            R.id.f18 -> ingresarValor("F", 175, 18)
            R.id.g18 -> ingresarValor("G", 176, 18)
            R.id.h18 -> ingresarValor("H", 177, 18)
            R.id.i18 -> ingresarValor("I", 178, 18)
            R.id.j18 -> ingresarValor("J", 179, 18)

            R.id.a19 -> ingresarValor("A", 180, 19)
            R.id.b19 -> ingresarValor("B", 181, 19)
            R.id.c19 -> ingresarValor("C", 182, 19)
            R.id.d19 -> ingresarValor("D", 183, 19)
            R.id.e19 -> ingresarValor("E", 184, 19)
            R.id.f19 -> ingresarValor("F", 185, 19)
            R.id.g19 -> ingresarValor("G", 186, 19)
            R.id.h19 -> ingresarValor("H", 187, 19)
            R.id.i19 -> ingresarValor("I", 188, 19)
            R.id.j19 -> ingresarValor("J", 189, 19)

            R.id.a20 -> ingresarValor("A", 190, 20)
            R.id.b20 -> ingresarValor("B", 191, 20)
            R.id.c20 -> ingresarValor("C", 192, 20)
            R.id.d20 -> ingresarValor("D", 193, 20)
            R.id.e20 -> ingresarValor("E", 194, 20)
            R.id.f20 -> ingresarValor("F", 195, 20)
            R.id.g20 -> ingresarValor("G", 196, 20)
            R.id.h20 -> ingresarValor("H", 197, 20)
            R.id.i20 -> ingresarValor("I", 198, 20)
            R.id.j20 -> ingresarValor("J", 199, 20)

            else -> {}
        }
    }

    //Para esta función lo que hace es que pide tres parámetros, el primero es para mostrar la columna, el cual es representada por una letra
    //El segundo es el número de casilla, para ir mostrando en cual casilla estas
    //Y el tercero es para indicar la fila en la que el usuario está
    //Estos parámetros los pide porque en esta función se despliega un AlertDialog, el cual
    //Sirve para editar el EditText correspondiente
    private fun ingresarValor(letraCasilla: String, numCasilla: Int, fila: Int) {
        //se crea el AlertDialog
        val builder = AlertDialog.Builder(this@MainActivity)
        //Se crea la ventana, el cual dirá la casilla en donde te encuentras
        val v =
            LayoutInflater.from(this@MainActivity).inflate(R.layout.item_dialog, null, false)
        builder.setTitle("Casilla $letraCasilla$fila")

        //Declaramos el EditText el cual es para editar el TextView
        val valor = v.findViewById<EditText>(R.id.etItem)
        //Aquí lo que hace es que si el TextView ya tiene algo escrito, en lugar de que no te muestre su contenido
        //Puedas editarlo
        valor.setText(datoCelda[numCasilla])

        builder.setView(v)
        //Aquí se crea un botón, el cual es donde actualiza el contenido del TextView
        builder.setPositiveButton("Actualizar") { _, _ ->
            textViewArray[numCasilla].text = funcion(valor.text.toString())
            datoCelda[numCasilla] = funcion(valor.text.toString())

        }
        //Aquí se crea un botón, el cual es para cancelar
        builder.setNegativeButton("Cancelar") { dialog, _ -> dialog.dismiss() }
        //Muestra la ventana
        builder.show()
    }

    private fun funcion(dato: String): String {
        try {
            var cadenaIngresada = dato.split("(").toTypedArray()
            var palabraClave =
                cadenaIngresada[0].uppercase()//Contendrá lo que se decidirá en el "switch"
            var valoresCortados = cadenaIngresada[1]
            var datosImportantes = valoresCortados.split(")").toTypedArray()
            var datosManipular = datosImportantes[0]//Contendrá los números o casillas

            if (palabraClave == "SUMA" || palabraClave == "SUM") {
                return suma(datosManipular)

            } else if (palabraClave == "PROMEDIO" || palabraClave == "AVERAGE") {
                return prom(datosManipular)

            } else if (palabraClave == "MAX") {
                return max(datosManipular)

            } else if (palabraClave == "MIN") {
                return min(datosManipular)
            }
            else if (palabraClave == "MODA" || palabraClave == "MODE") {
                return moda(datosManipular)
            }

        } catch (e: Exception) {

        }
        return dato
    }


    private fun suma(valor: String): String {
        try {
            var pr = "0"
            var cadenaIngresada = valor.split("(").toTypedArray()
            var palabraClave = cadenaIngresada[0].split(":").toTypedArray()
            val numbers = DoubleArray(palabraClave.size)
            var aux = 0.0

            for (i in (palabraClave.indices)) {
                try {
                    numbers[i] = palabraClave[i].toDouble()
                } catch (e: Exception) {
                    for (j in (0 until nombreCelda.size)) {
                        if (palabraClave[i] == nombreCelda[j]) {
                            try {
                                numbers[i] = textViewArray[j].text.toString().toDouble()
                            } catch (e: Exception) {
                                numbers[i] = 0.0
                            }

                        }
                    }
                }
                aux += numbers[i]
            }
            pr = String.format("%.2f", aux)
            return pr
        } catch (e: Exception) {
            e.printStackTrace()
            return "#¿NOMBRE?"
        }
    }

    private fun prom(valor: String): String {
        try {
            var pr = "0"
            var cadenaIngresada = valor.split("(").toTypedArray()
            var palabraClave = cadenaIngresada[0].split(":").toTypedArray()
            val numbers = DoubleArray(palabraClave.size)
            var aux = 0.0
            for (i in (palabraClave.indices)) {
                try {
                    numbers[i] = palabraClave[i].toDouble()
                } catch (e: Exception) {
                    for (j in (0 until nombreCelda.size)) {
                        if (palabraClave[i] == nombreCelda[j]) {
                            try {
                                numbers[i] = textViewArray[j].text.toString().toDouble()
                            } catch (e: Exception) {
                                numbers[i] = 0.0
                            }

                        }
                    }
                }

                aux += numbers[i]

            }
            aux /= palabraClave.size
            pr = String.format("%.2f", aux)
            return pr
        } catch (e: Exception) {
            e.printStackTrace()
            return "#¿NOMBRE?"
        }
    }

    private fun max(valor: String): String {
        try {
            var pr = "0"
            var cadenaIngresada = valor.split("(").toTypedArray()
            var palabraClave = cadenaIngresada[0].split(":").toTypedArray()
            val numbers = DoubleArray(palabraClave.size)

            for (i in (palabraClave.indices)) {
                try {
                    numbers[i] = palabraClave[i].toDouble()
                } catch (e: Exception) {
                    for (j in (0 until nombreCelda.size)) {
                        if (palabraClave[i] == nombreCelda[j]) {
                            try {
                                numbers[i] = textViewArray[j].text.toString().toDouble()
                            } catch (e: Exception) {
                                numbers[i] = 0.0
                            }

                        }
                    }
                }


            }
            val max = numbers.maxOrNull()
            pr = String.format("%.2f", max)
            return pr
        } catch (e: Exception) {
            e.printStackTrace()
            return "#¿NOMBRE?"
        }
    }


    private fun min(valor: String): String {
        try {
            var pr = "0"
            var cadenaIngresada = valor.split("(").toTypedArray()
            var palabraClave = cadenaIngresada[0].split(":").toTypedArray()
            val numbers = DoubleArray(palabraClave.size)
            for (i in (palabraClave.indices)) {
                try {
                    numbers[i] = palabraClave[i].toDouble()
                } catch (e: Exception) {
                    for (j in (0 until nombreCelda.size)) {
                        if (palabraClave[i] == nombreCelda[j]) {
                            try {
                                numbers[i] = textViewArray[j].text.toString().toDouble()
                            } catch (e: Exception) {
                                numbers[i] = 0.0
                            }

                        }
                    }
                }
            }
            val min = numbers.minOrNull()
            pr = String.format("%.2f", min)
            return pr
        } catch (e: Exception) {
            e.printStackTrace()
            return "#¿NOMBRE?"
        }
    }



    private fun moda(valor: String): String {
        try {
            var list = listOf<String>()
            var pr = "0"
            var cadenaIngresada = valor.split("(").toTypedArray()
            var palabraClave = cadenaIngresada[0].split(":").toTypedArray()
            var numbers = DoubleArray(palabraClave.size)
            for (i in (palabraClave.indices)) {
                try {
                    numbers[i] = palabraClave[i].toDouble()
                } catch (e: Exception) {
                    for (j in (0 until nombreCelda.size)) {
                        if (palabraClave[i] == nombreCelda[j]) {
                            try {
                                numbers[i] = textViewArray[j].text.toString().toDouble()
                                list += listOf(numbers[i].toString())

                            } catch (e: Exception) {
                                numbers[i] = 0.0
                                list += listOf(numbers[i].toString())
                            }

                        }
                    }
                }


            }

            var partNumber = list.groupingBy { it }.eachCount().filter { it.value > 1 }.toString().split("{")
            var onlyNumber = partNumber[1].split("}")
            return onlyNumber[0]
        } catch (e: Exception) {
            e.printStackTrace()
            return "#¿NOMBRE?"
        }
    }

    /* *
     *   Funciones especializadas en la obtanción de Permisos de USUARIO !!!!!
     *   Sacadas de algún lado de StackOverFlow...
     * */
    private fun askPermissionOnly() {
        askPermission(
            REQUEST_ID_WRITE_PERMISSION,
            Manifest.permission.WRITE_EXTERNAL_STORAGE
        )
        askPermission(
            REQUEST_ID_READ_PERMISSION,
            Manifest.permission.READ_EXTERNAL_STORAGE
        )
    }

    // With Android Level >= 23, you have to ask the user
    // for permission with device (For example read/write data on the device).
    private fun askPermission(requestId: Int, permissionName: String): Boolean {
        if (Build.VERSION.SDK_INT >= 23) {

            // Check if we have permission
            val permission = ActivityCompat.checkSelfPermission(this, permissionName)
            if (permission != PackageManager.PERMISSION_GRANTED) {
                // If don't have permission so prompt the user.
                requestPermissions(
                    arrayOf(permissionName),
                    requestId
                )
                return false
            }
        }
        return true
    }

    // When you have the request results
    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions!!, grantResults)
        // Note: If request is cancelled, the result arrays are empty.
        if (grantResults.isNotEmpty()) {
            when (requestCode) {
                REQUEST_ID_READ_PERMISSION -> {
                    run {
                        if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                            Toast.makeText(
                                applicationContext, "Permission Lectura Concedido!",
                                Toast.LENGTH_SHORT
                            ).show()
                        }
                    }
                    run {
                        if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                            //writeFile();
                            //
                            Toast.makeText(
                                applicationContext,
                                "Permission Escritura Concedido!",
                                Toast.LENGTH_SHORT
                            ).show()
                        }
                    }
                }
                REQUEST_ID_WRITE_PERMISSION -> {
                    if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                        Toast.makeText(
                            applicationContext,
                            "Permission Escritura Concedido!",
                            Toast.LENGTH_SHORT
                        ).show()
                    }
                }
            }
        } else {
            Toast.makeText(applicationContext, "Permission Cancelled!", Toast.LENGTH_SHORT).show()
        }

        // check condition
        if (requestCode == 1 && grantResults.isNotEmpty() && (grantResults[0]
                    == PackageManager.PERMISSION_GRANTED)
        ) {
            Toast.makeText(
                applicationContext,
                "Permission Escritura Concedido!",
                Toast.LENGTH_SHORT
            ).show()
        } else {
            // When permission is denied
            // Display toast
            Toast.makeText(applicationContext, "Permission Denied", Toast.LENGTH_SHORT).show()
        }
    }
}