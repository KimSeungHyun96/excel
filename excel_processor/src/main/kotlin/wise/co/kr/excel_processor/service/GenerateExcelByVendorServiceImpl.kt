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
    private val processKdicExcel: ProcessKdicExcel,
    private val processWDQPolExcel: ProcessWDQPolExcel,
    private val processWDQKosExcel: ProcessWDQKosExcel,
    private val processSDQEtcExcel: ProcessSDQEtcExcel,
    private val processSDQKeisExcel: ProcessSDQKeisExcel,
    private val processWDQKodExcel: ProcessWDQKodExcel,
    private val processEtcExcel: ProcessEtcExcel
) : GenerateExcelByVendorService {
    override fun generateExcelByVendor(
        sourceWorkbook: Workbook,
        targetWorkbook: Workbook,
        excelName: String
    ): Workbook {
        return when (getVendorName(sourceWorkbook)) {
            "WDQ" -> {
                processWDQExcel.generateWDQExcel(sourceWorkbook, targetWorkbook, excelName) //WDQ
            }
            "SDQ" -> {
                processSDQExcel.generateSDQExcel(sourceWorkbook,targetWorkbook, excelName) //SDQ
            }
            "SDQETC" -> {
                processSDQEtcExcel.generateSDQEtcExcel(sourceWorkbook,targetWorkbook, excelName) //다른양식의 SDQ
            }
            "DQMINER" -> {
                processDqminerExcel.generateDqminerExcel(sourceWorkbook,targetWorkbook, excelName) //DQMINER
            }
            "DQUBE" -> {
                processDqubeExcel.generateDqubeExcel(sourceWorkbook,targetWorkbook, excelName) //DQUBE
            }
            "DQ" -> {
                processDqExcel.generateDqExcel(sourceWorkbook,targetWorkbook, excelName) //DQ
            }
            "KDIC" -> {
                processKdicExcel.generateKdicExcel(sourceWorkbook,targetWorkbook, excelName) // 예금보험공사
            }
            "WDQPol" -> {
                processWDQPolExcel.generateWDQPolExcel(sourceWorkbook,targetWorkbook, excelName) //경찰청
            }
            "WDQKos" -> {
                processWDQKosExcel.generateWDQKosExcel(sourceWorkbook,targetWorkbook, excelName) //산업안전보건공단
            }
            "SDQKeis" -> {
                processSDQKeisExcel.generateSDQKeisExcel(sourceWorkbook,targetWorkbook, excelName) //한국고용정보원
            }
            "WDQKod" -> {
                processWDQKodExcel.generateWDQKodExcel(sourceWorkbook,targetWorkbook, excelName) //신용보증기금
            }
            "ETC" -> {
                println("22222")
                processEtcExcel.generateEtcExcel(sourceWorkbook,targetWorkbook, excelName) // 기타양식
            }
            else -> {
                throw IllegalArgumentException("vendor not found")
            }
        }

    }
    //진단 도구 판별 함수
    private fun getVendorName(sourceWorkbook: Workbook): String {

        val sheet = sourceWorkbook.getSheet("(진단결과)값진단결과")
        val sheet2 = sourceWorkbook.getSheet("값진단결과") // 한국고용정보원
        val sheet3 = sourceWorkbook.getSheet("진단현황") // 한국고용정보원

        // sheet1 관련 셀
        var cell = sheet?.getRow(0)?.getCell(1)
        var cellEtc = sheet?.getRow(1)?.getCell(1)
        var cellSDQ = sheet?.getRow(1)?.getCell(0) // 다른양식의 SDQ
        var cell_Organization = sheet?.getRow(3)?.getCell(2) // 경찰청, 한국산업안전보건공단, 한국산업은행
        var cell_Organization2 = sheet?.getRow(6)?.getCell(2) // 예금보험공사

        // sheet2 관련 셀
        var cell2 = sheet2?.getRow(0)?.getCell(1)
        var cell_Organization3 = sheet2?.getRow(4)?.getCell(2) // 한국고용정보원

        if (cell != null && cell.stringCellValue.isNotBlank()) {
            return when {
                // 순서변경 X (경찰청 1순위)
                cell.stringCellValue.equals("WISE DQ 값진단 결과 보고서") && cell_Organization?.stringCellValue.equals("경찰청") -> "WDQPol" //경찰청
                cell.stringCellValue.equals("값진단 결과 보고서") && cell_Organization?.stringCellValue.equals("한국산업안전보건공단") -> "WDQKos" //산업안전보건공단
                cell.stringCellValue.contains("WISE") -> "WDQ"
                cell.stringCellValue.contains("SDQ") -> "SDQ"
                cell.stringCellValue.contains("DQube") -> "DQUBE"
                cell.stringCellValue.equals("값진단 결과 보고서") && cellEtc?.stringCellValue?.contains("DQMINER") == true -> "DQMINER"
                cell.stringCellValue.equals("DQ 값 진단 종합 현황") && cell_Organization?.stringCellValue.equals("한국산업은행") -> "DQ"
                cell.stringCellValue.equals("신용보증기금 자체품질관리시스템 값진단 결과 보고서") -> "WDQKod" //신용보증기금
                else -> throw IllegalArgumentException("cell vendor not found")
            }
        } else if (cellEtc != null && cellEtc.stringCellValue.isNotBlank()) {
            return when {
                cellEtc.stringCellValue.equals("값 진단 종합 현황") && cell_Organization2?.stringCellValue.equals("예금보험공사") -> "KDIC"
                else -> throw IllegalArgumentException("cellEtc vendor not found")
            }
        } else if ((cellEtc == null || cellEtc.stringCellValue.isBlank()) && cellSDQ != null && cellSDQ.stringCellValue.isNotBlank()) {
            // cellEtc가 공백이거나 null일 때 cellSDQ 분기로 넘어감
            return when {
                cellSDQ.stringCellValue.contains("SDQ") -> "SDQETC"
                else -> throw IllegalArgumentException("cellSDQ vendor not found")
            }
        } else if (cell2 != null && cell2.stringCellValue.equals("값진단 결과 보고서") && cell_Organization3?.stringCellValue.equals("한국고용정보원")) {
            // sheet2에서 값 확인 후 SDQKeis로 분기
            return when {
                cell2.stringCellValue.equals("값진단 결과 보고서") -> "SDQKeis" //한국고용정보원
                else -> throw IllegalArgumentException("cellSDQ vendor not found")
            }
        } else {
            throw IllegalArgumentException("vendor not found")
        }
    }
}

