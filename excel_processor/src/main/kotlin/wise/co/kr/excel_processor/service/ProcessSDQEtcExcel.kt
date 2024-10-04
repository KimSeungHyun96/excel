package wise.co.kr.excel_processor.service

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.springframework.stereotype.Service
import java.text.DecimalFormat
import java.time.format.DateTimeFormatter


@Service
class ProcessSDQEtcExcel : ProcessExcelService {
    fun generateSDQEtcExcel(
        sourceWorkbook: Workbook, targetWorkbook: Workbook, excelName: String
    ): Workbook {
        // 값진단결과
        val sourceSheet0 = sourceWorkbook.getSheet("(진단결과)값진단결과") ?: throw IllegalArgumentException("Sheet 0 not found")
        // 진단대상테이블
//        val sourceSheet1 =
//            sourceWorkbook.getSheet("(테이블선정)진단대상테이블") ?: throw IllegalArgumentException("Sheet 1 not found")
        // 도메인
//        val sourceSheet2 = sourceWorkbook.getSheet("(룰설정)도메인") ?: throw IllegalArgumentException("Sheet 2 not found")
        // 업무규칙
        val sourceSheet4 = sourceWorkbook.getSheet("(룰설정)업무규칙") ?: throw IllegalArgumentException("Sheet 4 not found")


        val dataHashMap0 = generateHashMap(sourceSheet0)
        dataHashMap0["파일명"] = excelName
        dataHashMap0["진단도구명"] = sourceSheet0.getRow(1).getCell(0)
        //빈칸이 존재하는 셀이 있다면, 해당 셀이 마지막로우이기 때문에 건수가 늘어갈 가능성 존재
        var cursor = 0
        val totalRows = sourceSheet4.physicalNumberOfRows // 전체 행의 개수

        // 실제 데이터를 가진 행을 찾는 반복문
        for (i in 0 until totalRows) {
            val row = sourceSheet4.getRow(i)
            // 각 행의 첫 번째 셀이 null이 아니고 공백이 아닐 때만 cursor를 증가
            if (row != null && !row.getCell(0).toString().isNullOrBlank()) {
                cursor++
            }
        }
        // 업무규칙 수를 빈 셀을 제외한 실제 데이터 행의 수로 설정
        dataHashMap0["업무규칙 수"] = (cursor-1).toString()
//        dataHashMap0["업무규칙 수"] = (sourceSheet4.physicalNumberOfRows - 1).toString()
        dataHashMap0["작업시간"] = getCurrentKoreanTime()

        val targetSheet0 = targetWorkbook.getSheetAt(0) ?: throw IllegalArgumentException("Target Sheet 0 not found")
        val headerRowSheet0 = targetSheet0.getRow(0) ?: targetSheet0.createRow(0)
        val newRowSheet0 = targetSheet0.createRow(targetSheet0.lastRowNum + 1)
        val colNumSheet0 = headerRowSheet0.physicalNumberOfCells

        for (i in 0 until colNumSheet0) {
            val headerCell = headerRowSheet0.getCell(i) ?: continue
            val headerCellValue = headerCell.stringCellValue
            val value = dataHashMap0[headerCellValue]?.toString() ?: ""
            newRowSheet0.createCell(i).setCellValue(value)
        }

        // Sheet 1 처리
        val targetSheet1 = targetWorkbook.getSheetAt(1) ?: throw IllegalArgumentException("Target Sheet 1 not found")
        val headerRowSheet1 = targetSheet1.getRow(0) ?: targetSheet1.createRow(0)
        val rowNumSheet1 = targetSheet1.physicalNumberOfRows
        val colNumSheet1 = headerRowSheet1.physicalNumberOfCells

        //qualityIndicatorNameList 는 key 가 품질지표명이고 value 가 진단건수, 오류건수, 오류율을 갖고있는 List<Stirng> 인 HashMap의 List
        //
        val qualityIndicatorNameList = dataHashMap0["품질지표명"] as MutableList<*>

        for (i in rowNumSheet1 until rowNumSheet1 + qualityIndicatorNameList.size) {
            val newRowSheet1 = targetSheet1.createRow(i)
            val qualityHashMap = qualityIndicatorNameList[i - rowNumSheet1] as HashMap<String, List<String>>
            val key = qualityHashMap.keys.first().toString()
            val valueList = qualityHashMap.getValue(key)

            for (j in 0 until colNumSheet1) {
                val headerCell = headerRowSheet1.getCell(j) ?: continue
                when (val headerCellValue = headerCell.stringCellValue) {
                    "품질지표명" -> newRowSheet1.createCell(j).setCellValue(key)
                    "진단건수" -> newRowSheet1.createCell(j).setCellValue(valueList[0])
                    "오류건수" -> newRowSheet1.createCell(j).setCellValue(valueList[1])
                    "오류율" -> newRowSheet1.createCell(j).setCellValue(valueList[2])
                    else -> {
                        val value = dataHashMap0[headerCellValue]?.toString() ?: ""
                        newRowSheet1.createCell(j).setCellValue(value)
                    }
                }
            }

        }
        //sheet 2
        //상태가 "대상"인 row 의 테이블명, 상태, 범위조건, 의견 가져오기
        //별도의 리스트는 필요하지 않을 듯

//        val targetSheet2 = targetWorkbook.getSheetAt(2) ?: throw IllegalArgumentException("Target Sheet 2 not found")
//        val targetHeaderRow = targetSheet2.getRow(0) ?: throw IllegalArgumentException("TargetHeaderRow is not created")
//        for (i in 0 until sourceSheet1.physicalNumberOfRows) {
//            val sourceHeaderRow = sourceSheet1.getRow(0)
//            val sourceRow = sourceSheet1.getRow(i) ?: continue
//            val statusCell = sourceRow.getCell(1)
//
//            if (statusCell.stringCellValue == "대상") {
//                val targetRow = targetSheet2.createRow(targetSheet2.physicalNumberOfRows)
//                for (j in 0 until targetHeaderRow.physicalNumberOfCells) {
//                    if (targetHeaderRow.getCell(j).toString() == "DBMS명") {
//                        val value = dataHashMap0["DBMS명"].toString()
//                        targetRow.createCell(j).setCellValue(value)
//                    } else if (targetHeaderRow.getCell(j).toString() == "스키마명") {
//                        val value = dataHashMap0["DB서비스명"].toString()
//                        targetRow.createCell(j).setCellValue(value)
//                    } else if (targetHeaderRow.getCell(j).toString() == "테이블명") {
//                        val sourceCell = sourceRow.getCell(j - 2)
//                        val value = getCellValueAsString(sourceCell)
//                        targetRow.createCell(j).setCellValue(value)
//                    } else {
//                        val sourceCell = sourceRow.getCell(j - 2)
//                        val value = getCellValueAsString(sourceCell)
//                        targetRow.createCell(j).setCellValue(value)
//                    }
//                }
//            } else {
//                continue
//            }
//        }

        //sheet 3
        //검증룰 명이 존재하는 row 의 테이블명, 컬럼명, 데이터타입, 검증룰명,품질지표명, 검증룰, 오류제외데이터, 의견 가져오기
        //source sheet 에서 row 를 기준으로 반복문을 돌면서 일치하는 단어가 있는 헤더 기준 셀을 복사해서 value 로 집어넣기
//        val targetSheet3 = targetWorkbook.getSheetAt(3) ?: throw IllegalArgumentException("Target Sheet 2 not found")
//
//        for (i in 1 until sourceSheet2.physicalNumberOfRows) {
//            val sourceHeaderRow = sourceSheet2.getRow(0) ?: throw IllegalArgumentException("HeaderRow is not created")
//            val targetSourceHeaderRow = targetWorkbook.getSheetAt(3).getRow(0)
//            val sourceRow = sourceSheet2.getRow(i) ?: continue
//            val elementHashMap: HashMap<String, String> = hashMapOf()
//            //SDQ 도메인 순서: 테이블, 컬럼, 데이터타입, 진단기준명, 도메인, 진단기준, 오류제외데이터
//            if (sourceRow.getCell(5).toString().isNotBlank()) {
//                for (j in 0 until sourceHeaderRow.physicalNumberOfCells) {
//                    val sourceHeaderCell = sourceHeaderRow.getCell(j).toString()
//                    val sourceCell = sourceRow.getCell(j).toString()
//
//                    elementHashMap["DBMS명"] = dataHashMap0["DBMS명"].toString()
//                    elementHashMap["스키마명"] = dataHashMap0["DB서비스명"].toString()
//                    if (sourceHeaderCell == "테이블") {
//                        elementHashMap["테이블명"] = sourceCell
//                    } else if (sourceHeaderCell == "컬럼") {
//                        elementHashMap["컬럼명"] = sourceCell
//                    } else if (sourceHeaderCell == "데이터타입") {
//                        elementHashMap["데이터타입"] = sourceCell
//                    } else if (sourceHeaderCell == "진단기준명") {
//                        elementHashMap["검증룰명"] = sourceCell
//                    } else if (sourceHeaderCell == "도메인") {
//                        elementHashMap["품질지표명"] = sourceCell
//                    } else if (sourceHeaderCell == "진단기준") {
//                        elementHashMap["검증룰"] = sourceCell
//                    } else if (sourceHeaderCell == "오류제외데이터") {
//                        elementHashMap["오류제외데이터"] = sourceCell
//                    } else if (sourceHeaderCell.contains("의견")) {
//                        elementHashMap["의견"] = sourceCell
//                    }
//
//                }
//
//                val targetRow = targetSheet3.createRow(targetSheet3.physicalNumberOfRows)
//
//                for (k in 0 until targetSourceHeaderRow.physicalNumberOfCells) {
//                    targetRow
//                        .createCell(k).setCellValue(
//                            elementHashMap.getValue(
//                                targetSheet3.getRow(0).getCell(k).toString()
//                            )
//                        )
//
//                }
//
//                elementHashMap.clear()
//
//            } else {
//                continue
//            }
//        }
        return targetWorkbook

    }


    private fun generateHashMap(sourceSheet: Sheet): HashMap<String, Any> {
        val resultMap = HashMap<String, Any>()
        val keywordsToRow = listOf("기관명", "시스템", "DB명", "DB서비스명", "DB종류", "버전", "출력일 :")
        val keywordsToCol = listOf("진단건수", "오류건수", "오류율(%)")

        for (rowIndex in 0..sourceSheet.lastRowNum) {
            val row = sourceSheet.getRow(rowIndex) ?: continue
            for (cellIndex in 0..row.lastCellNum) {
                val cell = row.getCell(cellIndex) ?: continue
                val cellValue = getCellValueAsString(cell)
                val qualityIndicatorHashMapList: MutableList<HashMap<String, List<String>>> = mutableListOf()

                when {
                    keywordsToRow.contains(cellValue) -> {
                        val nextCell = row.getCell(cellIndex + 1)
                        if (nextCell != null) {
                            when (cellValue) {
                                "시스템" -> {
                                    resultMap["정보시스템명"] = getCellValueAsString(nextCell)
                                    resultMap["시스템명"] = getCellValueAsString(nextCell)
                                }

                                "DB명" -> {
                                    resultMap["DBMS명"] = getCellValueAsString(nextCell)
                                    resultMap["DB명"] = getCellValueAsString(nextCell)
                                }

                                "DB서비스명" -> {
                                    resultMap["DB서비스명"] = getCellValueAsString(nextCell)
                                    resultMap["DB명"] = getCellValueAsString(nextCell)
                                }

                                "DB종류" -> {
                                    resultMap["DBMS종류"] = getCellValueAsString(nextCell)
                                    resultMap["DB종류"] = getCellValueAsString(nextCell)
                                }

                                "버전" -> {
                                    resultMap["버전"] = getCellValueAsString(nextCell)
                                }

                                "출력일 :" -> {
                                    resultMap["출력일"] = getCellValueAsString(nextCell)
                                }

                                else -> {
                                    resultMap[cellValue] = getCellValueAsString(nextCell)
                                }
                            }


                        }
                    }

                    keywordsToCol.contains(cellValue) -> {
                        for (i in sourceSheet.lastRowNum downTo rowIndex) {
                            val lastCell = sourceSheet.getRow(i).getCell(cellIndex)
                            if (lastCell != null && getCellValueAsString(lastCell).isNotBlank()) {
                                when (cellValue) {
                                    "진단건수" -> {
                                        resultMap["총 진단건수"] = getCellValueAsString(lastCell)
                                    }

                                    "오류건수" -> {
                                        resultMap["총 오류건수"] = getCellValueAsString(lastCell)
                                    }

                                    "오류율(%)" -> {
                                        resultMap["총 오류율"] = getCellValueAsString(lastCell)+"%"
                                    }
                                }
                                break
                            }
                        }
                    }

//                    cellValue.contains("출력일") -> {
//                        resultMap["출력일"] = cellValue.substring(cellValue.indexOf(":") + 1).trim()
//                    }

                    cellValue == "진단항목" -> {
                        var newRowIndex = rowIndex + 1
                        var newRow = sourceSheet.getRow(newRowIndex)
                        var newCellValue = getCellValueAsString(newRow.getCell(cellIndex))

                        while (newCellValue.isNotBlank()) {
                            val qualityIndicatorHashMap: HashMap<String, List<String>> = hashMapOf()

                            val diagnosisCount = getCellValueAsString(newRow.getCell(cellIndex + 2))
                            val errorCount = getCellValueAsString(newRow.getCell(cellIndex + 3))
                            var errorRate = getCellValueAsString(newRow.getCell(cellIndex + 4))+"%"

                            // errorRate의 값이 ".123445%"라면 "0.123445%"로 변환
                            if (errorRate.startsWith(".")) {
                                errorRate = "0"+errorRate // "."으로 시작하면 앞에 "0"을 붙임
                            }

                            qualityIndicatorHashMap[newCellValue] = listOf(diagnosisCount, errorCount, errorRate)
                            qualityIndicatorHashMapList.add(qualityIndicatorHashMap)

                            // 다음 행으로 이동
                            newRowIndex++
                            newRow = sourceSheet.getRow(newRowIndex)
                            newCellValue = getCellValueAsString(newRow.getCell(cellIndex))
                        }
                        resultMap["품질지표명"] = qualityIndicatorHashMapList
                    }

                    else -> {
                        continue
                    }
                }
            }
        }
        return resultMap
    }

    private fun getCellValueAsString(cell: Cell): String {
        val decimalFormat = DecimalFormat("#,###") // 쉼표로 자리 구분

        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.localDateTimeCellValue.toString()
                } else {
                    decimalFormat.format(cell.numericCellValue) // 숫자를 쉼표로 구분된 형태로 출력
                }
            }
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            CellType.FORMULA -> {
                // 수식 셀의 계산된 값을 직접 가져옵니다.
                when (cell.cachedFormulaResultType) {
                    CellType.STRING -> cell.stringCellValue
                    CellType.NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            cell.localDateTimeCellValue.format(DateTimeFormatter.ISO_LOCAL_DATE_TIME)
                        } else {
                            decimalFormat.format(cell.numericCellValue) // 숫자를 쉼표로 구분된 형태로 출력
                        }
                    }
                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    else -> cell.stringCellValue // 기타 경우 문자열로 처리
                }
            }
            else -> ""
        }
    }


}