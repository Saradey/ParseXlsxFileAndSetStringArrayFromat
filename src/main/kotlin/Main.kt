import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileWriter
import java.io.IOException

/**
 * First arg path xlsx file
 * Second arg sheet index
 * Third arg first cell
 * Fourth arg second cell
 * Fifth arg path save xml file <string-array>First-Cell|Second-Cell</string-array>
 */
fun main(args: Array<String>) {
    val work = XSSFWorkbook(args.first())
    val sheet = work.getSheetAt(args[1].toInt())
    val rows = sheet.iterator()
    val stringScan = StringBuilder()
    stringScan.append("<string-array name=\"thisYourName\">")
    stringScan.append("\n")
    rows.forEach { row ->
        stringScan.append("\t<item>${row.getCell(args[2].toInt())}|${row.getCell(args[3].toInt())}</item>")
        stringScan.append("\n")
    }
    stringScan.append("</string-array>")

    try {
        val fileWriter = FileWriter(args[4], false)
        fileWriter.write(stringScan.toString())
        fileWriter.flush()
    } catch (exp: IOException) {
        exp.printStackTrace()
    } finally {
        println("Enjoy!")
    }
}