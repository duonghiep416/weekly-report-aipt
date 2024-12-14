const ExcelJS = require('exceljs')
const moment = require('moment')
const path = require('path')
const fs = require('fs')

// CẤU HÌNH CHÍNH
const CONFIG = {
  inputFilePath: path.join(__dirname, 'daily-report.xlsx'), // File báo cáo hàng ngày
  templatePath: path.join(__dirname, 'template.xlsx'), // File template
  outputDirectory: path.join(__dirname, 'output'), // Thư mục lưu file báo cáo tuần
  outputFileName: 'BaoCaoTuan', // Tên file báo cáo tuần
  sheetName: '122024', // Tên sheet trong file đầu vào (ví dụ: 122024)
  weekInfo: {
    name: 'Dương Đức Hiệp', // Tên nhân sự
    startDate: '2024-12-09', // Ngày bắt đầu tuần (YYYY-MM-DD)
    endDate: '2024-12-15', // Ngày kết thúc tuần (YYYY-MM-DD)
    month: 12, // Tháng
    year: 2024 // Năm
  }
}

// HÀM ĐỌC DỮ LIỆU TỪ FILE
async function readSheetData(
  filePath,
  sheetName,
  startDate,
  endDate,
  userName
) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  // Lấy sheet theo tên
  const worksheet = workbook.getWorksheet(sheetName)
  if (!worksheet) {
    throw new Error(`Không tìm thấy sheet với tên: ${sheetName}`)
  }

  // Xử lý tháng và năm từ tên sheet
  const month = parseInt(sheetName.slice(0, 2), 10) // Lấy 2 ký tự đầu tiên làm tháng
  const year = parseInt(sheetName.slice(2), 10) // Lấy 4 ký tự sau làm năm

  // Lọc dữ liệu từ cột A (ngày), B (tên), C (nội dung)
  const reports = []
  worksheet.eachRow((row, rowIndex) => {
    if (rowIndex > 1) {
      // Bỏ qua tiêu đề
      const day = parseInt(row.getCell(1).value, 10) // Ngày trong cột A
      const name = row.getCell(2).value // Tên trong cột B
      const task = row.getCell(3).value // Công việc trong cột C

      // Tạo ngày đầy đủ (YYYY-MM-DD) từ ngày, tháng, năm
      const fullDate = moment(`${year}-${month}-${day}`, 'YYYY-MM-DD')

      // Lọc theo ngày và tên chính xác
      if (
        fullDate.isBetween(startDate, endDate, undefined, '[]') &&
        name === userName
      ) {
        reports.push({ date: fullDate.format('YYYY-MM-DD'), name, task })
      }
    }
  })

  return reports
}

// HÀM TẠO BÁO CÁO TUẦN
async function generateWeeklyReport(
  reports,
  templatePath,
  outputPath,
  weekInfo
) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templatePath)

  // Chọn sheet đầu tiên
  const worksheet = workbook.getWorksheet('Form')

  // Kiểm tra nếu worksheet không tồn tại
  if (!worksheet) {
    throw new Error('Sheet đầu tiên trong file template không tồn tại.')
  }

  // Lọc ra các công việc có dữ liệu hợp lệ
  const validReports = reports.filter((report) => report.task && report.date)

  // Tách từng công việc riêng lẻ từ các báo cáo có nhiều dòng
  const expandedReports = []
  validReports.forEach((report) => {
    const tasks = report.task
      .split('\n') // Tách từng dòng công việc
      .map((task) => task.trim()) // Loại bỏ khoảng trắng thừa
      .filter((task) => task) // Loại bỏ dòng trống
    tasks.forEach((task) => {
      expandedReports.push({ date: report.date, name: report.name, task })
    })
  })

  // Tìm tất cả dự án duy nhất từ các task
  const uniqueProjects = new Set()
  expandedReports.forEach((report) => {
    const match = report.task.match(/: (.+)$/) // Tìm tên dự án sau dấu ":"
    if (match) {
      uniqueProjects.add(match[1].trim())
    }
  })

  // Điền số lượng phần mềm tham gia vào ô C8
  worksheet.getCell('C8').value = uniqueProjects.size

  // Điền dữ liệu vào file báo cáo
  let currentRow = 9 // Dòng bắt đầu ghi dữ liệu
  uniqueProjects.forEach((project) => {
    // Ghi tên dự án (Cột C)
    worksheet.getCell(`C${currentRow}`).value = project
    currentRow++

    // Ghi công việc của dự án (Cột D và E)
    expandedReports
      .filter((report) => {
        const match = report.task.match(/: (.+)$/) // Tìm tên dự án
        return match && match[1].trim() === project
      })
      .forEach((report) => {
        const taskContent = report.task.split(':')[0].trim() // Lấy phần nội dung công việc trước dấu ":"
        worksheet.getCell(`D${currentRow}`).value = `${taskContent}` // Nội dung công việc
        worksheet.getCell(`E${currentRow}`).value = report.date // Ngày
        currentRow++
      })
  })

  // Lưu file báo cáo
  await workbook.xlsx.writeFile(outputPath)
}

// HÀM CHÍNH
async function main() {
  const { inputFilePath, templatePath, outputDirectory, sheetName, weekInfo } =
    CONFIG

  // Đảm bảo thư mục output tồn tại
  if (!fs.existsSync(outputDirectory)) {
    fs.mkdirSync(outputDirectory)
  }

  // Thêm timestamp vào tên file
  const timestamp = moment().format('YYYYMMDD_HHmmss') // Định dạng thời gian: YYYYMMDD_HHmmss
  const outputFileName = `${CONFIG.outputFileName}_${timestamp}.xlsx` // Tên file với timestamp
  const outputFilePath = path.join(outputDirectory, outputFileName)

  // Đọc dữ liệu từ sheet
  const reports = await readSheetData(
    inputFilePath,
    sheetName,
    weekInfo.startDate,
    weekInfo.endDate,
    weekInfo.name
  )

  // Tạo báo cáo tuần
  await generateWeeklyReport(reports, templatePath, outputFilePath, weekInfo)

  console.log(`File báo cáo tuần đã tạo tại: ${outputFilePath}`)
}

// Chạy chương trình
main().catch((err) => console.error(err))
