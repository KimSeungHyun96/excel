package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service

@Service
class GenerateExcelByVendorServiceImpl(
    private val processWDQExcel: ProcessWDQExcel,
    private val processSDQExcel: ProcessSDQExcel,
    private val processDqminerExcel: ProcessDqminerExcel,
    private val processDqubeExcel: ProcessDqubeExcel,
    private val processDqExcel: ProcessDqExcel,
    private val processEtcExcel: ProcessEtcExcel,
    private val processWDQPolExcel: ProcessWDQPolExcel,
    private val processSDQEtcExcel: ProcessSDQEtcExcel
) : GenerateExcelByVendorService {
    override fun generateExcelByVendor(
        sourceWorkbook: Workbook,
        targetWorkbook: Workbook,
        excelName: String
    ): Workbook {
        return when (getVendorName(sourceWorkbook)) {
            "WDQ" -> {
                processWDQExcel.generateWDQExcel(sourceWorkbook, targetWorkbook, excelName)
            }
            "SDQ" -> {
                processSDQExcel.generateSDQExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "SDQETC" -> {
                processSDQEtcExcel.generateSDQEtcExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "DQMINER" -> {
                processDqminerExcel.generateDqminerExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "DQUBE" -> {
                processDqubeExcel.generateDqubeExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "DQ" -> {
                processDqExcel.generateDqExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "ETC" -> {
                processEtcExcel.generateEtcExcel(sourceWorkbook,targetWorkbook, excelName)
            }
            "WDQPol" -> {
                processWDQPolExcel.generateWDQPolExcel(sourceWorkbook,targetWorkbook, excelName)
            }

            else -> {
                throw IllegalArgumentException("vendor not found")
            }
        }

    }
    //진단 도구 판별 함수
    private fun getVendorName(sourceWorkbook: Workbook): String {
//

//        val sheet = sourceWorkbook.getSheet("(진단결과)값진단결과")
//
//        var vendor :String = ""
//
//        if(sheet.getRow(0).getCell(1).toString().isNotEmpty()){
//            vendor = "first"
//        }else if(sheet.getRow(1).getCell(1).toString().isNotEmpty()){
//            vendor = "second"
//        }else if(sheet.getRow(1).getCell(0).toString().isNotEmpty()){
//            vendor = "third"
//        }
//
//        val cell = sheet.getRow(0).getCell(1)
//        val cellEtc = sheet.getRow(1).getCell(1)
//        val cellSDQ = sheet.getRow(1).getCell(0) // 다른양식의 SDQ
//
//        if(vendor == "first") {
//            return when {
//                cell.stringCellValue.equals("WISE DQ 값진단 결과 보고서") -> "WDQPol"
//                cell.stringCellValue.equals("DQ 값 진단 종합 현황") -> "DQ"
//                cell.stringCellValue.contains("WISE") -> "WDQ"
//                cell.stringCellValue.contains("SDQ") -> "SDQ"
//                cell.stringCellValue.contains("DQube") -> "DQUBE"
//                cellEtc.stringCellValue.equals("값진단 결과 보고서") -> "DQMINER"
//
//                else -> throw IllegalArgumentException("vendor not found")
//            }
//        } else if(vendor == "second") {
//            return when{
//                cellEtc.stringCellValue.equals("값 진단 종합 현황") -> "ETC"
//                else -> throw IllegalArgumentException("vendor not found")
//            }
//        } else if(vendor == "third"){
//            return when{
//                cellSDQ.stringCellValue.equals("SDQ 값진단 결과 보고서") -> "SDQETC"
//
//                else -> throw IllegalArgumentException("vendor not found")
//            }
//        } else {
//            throw IllegalArgumentException("vendor not found")
//        }


        val sheet = sourceWorkbook.getSheet("(진단결과)값진단결과")
        val cell = sheet.getRow(0).getCell(1)
        val cellEtc = sheet.getRow(1).getCell(1)
        val cellSDQ = sheet.getRow(1).getCell(0) // 다른양식의 SDQ

        return when {

            // 순서변경 X (경찰청 1순위)
            cell.stringCellValue.equals("WISE DQ 값진단 결과 보고서") -> "WDQPol"
            cell.stringCellValue.contains("WISE") -> "WDQ"
            cell.stringCellValue.contains("SDQ") -> "SDQ"
            cell.stringCellValue.contains("DQube") -> "DQUBE"
            cellEtc.stringCellValue.contains("DQMINER") -> "DQMINER"
            cell.stringCellValue.equals("DQ 값 진단 종합 현황") -> "DQ"

            // exception 처리 나온 항목 대상으로 ETC 및 SDQETC 진행
//            cellEtc.stringCellValue.equals("값 진단 종합 현황") -> "ETC"
//            cellSDQ.stringCellValue.contains("SDQ") -> "SDQETC"

            else -> throw IllegalArgumentException("vendor not found")
        }
    }
}

