// =================================================================
//                    CONFIGURATION VARIABLES
// =================================================================
// --- ID ของ Google Sheet ที่ใช้เป็นฐานข้อมูล ---
const SPREADSHEET_ID = '1FCc092xqiSbQlq-5dDFGosUA8IKbk2lh7q5VIPj0EOM';
// --- ชื่อชีตสำหรับเก็บข้อมูลใบงาน ---
const SHEET_NAME = 'InstallationPlan';
// --- ชื่อชีตสำหรับเก็บข้อมูลวันหยุด ---
const HOLIDAY_SHEET_NAME = 'Holidays';
// --- ชื่อชีตสำหรับเก็บข้อมูลพนักงาน ---
const EMPLOYEE_SHEET_NAME = 'Employees';
// --- ID ของโฟลเดอร์หลักใน Google Drive สำหรับเก็บรูปภาพและไฟล์แนบ ---
const FOLDER_ID = '1mLFLR0Jq9Cwu1nUfXa35XsTVytxuDqDR';


// =================================================================
//                      MAIN APP FUNCTIONS
// =================================================================

/**
 * ฟังก์ชันหลักที่ทำงานเมื่อมีการเรียก URL ของเว็บแอป
 * @param {object} e - อ็อบเจ็กต์ Event จาก Google Apps Script
 * @returns {HtmlOutput} - หน้าเว็บ HTML ที่จะแสดงผล
 */
function doGet(e) {
    try {
        // สร้าง HTML output จากไฟล์ Index.html โดยตรง
        return HtmlService.createHtmlOutputFromFile('Index')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setTitle('ระบบแผนงานติดตั้งและบริการ');
    } catch (error) {
        Logger.log('FATAL ERROR in doGet: ' + error.stack);
        return HtmlService.createHtmlOutput(`<h1>เกิดข้อผิดพลาดร้ายแรง</h1><p>ไม่สามารถโหลดแอปพลิเคชันได้ โปรดติดต่อผู้ดูแลระบบ</p><p>Error: ${error.message}</p>`);
    }
}

/**
 * ดึงข้อมูลเริ่มต้นทั้งหมดที่จำเป็นสำหรับแอปพลิเคชัน
 * @returns {object} - อ็อบเจ็กต์ที่ประกอบด้วยสถานะและข้อมูล (ใบงาน, วันหยุด, พนักงาน, URL ของเว็บแอป)
 */
function getInitialData() {
    try {
        const repairDataResult = getRepairData();
        if (repairDataResult.status === 'error') return repairDataResult;

        const holidayDataResult = getHolidayData();
        if (holidayDataResult.status === 'error') Logger.log(`Warning: Could not retrieve holiday data. ${holidayDataResult.message}`);

        const employeeDataResult = getEmployeeData();
        if (employeeDataResult.status === 'error') Logger.log(`Warning: Could not retrieve employee data. ${employeeDataResult.message}`);

        const webAppUrl = ScriptApp.getService().getUrl();

        return {
            status: 'success',
            data: {
                repairData: repairDataResult.data || [],
                holidayData: holidayDataResult.data || [],
                employeeData: employeeDataResult.data || [],
                webAppUrl: webAppUrl
            }
        };
    } catch (error) {
        Logger.log('ERROR in getInitialData: ' + error.stack);
        return {
            status: 'error',
            message: `Failed to load initial data: ${error.message}`
        };
    }
}

// =================================================================
//                      DATA RETRIEVAL FUNCTIONS
// =================================================================

/**
 * ดึงข้อมูลใบงานทั้งหมดจาก Google Sheet
 * @returns {object} - อ็อบเจ็กต์พร้อมข้อมูลใบงานทั้งหมด
 */
function getRepairData() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

        const values = sheet.getDataRange().getDisplayValues();
        if (values.length < 2) return {
            status: 'success',
            data: []
        };

        const headers = values[0].map(h => h.trim());
        const data = values.slice(1).map(row => {
            return headers.reduce((obj, header, index) => {
                if (header) obj[header] = row[index];
                return obj;
            }, {});
        }).filter(obj => obj['เลขที่ใบสั่งงาน']);

        data.sort((a, b) => (new Date(b['วันที่แจ้ง']).getTime() || 0) - (new Date(a['วันที่แจ้ง']).getTime() || 0));

        return {
            status: 'success',
            data: data
        };
    } catch (error) {
        Logger.log('ERROR in getRepairData: ' + error.stack);
        return {
            status: 'error',
            message: `Failed to get repair data: ${error.message}`
        };
    }
}

/**
 * ดึงข้อมูลวันหยุดจาก Google Sheet
 * @returns {object} - อ็อบเจ็กต์พร้อมข้อมูลวันหยุด
 */
function getHolidayData() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HOLIDAY_SHEET_NAME);
        if (!sheet || sheet.getLastRow() < 2) return {
            status: 'success',
            data: []
        };

        const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        const holidays = values.map(row => {
            const date = new Date(row[0]);
            return !isNaN(date.getTime()) && row[1] ? {
                date: Utilities.formatDate(date, "UTC", "yyyy-MM-dd"),
                title: row[1]
            } : null;
        }).filter(Boolean);

        return {
            status: 'success',
            data: holidays
        };
    } catch (error) {
        Logger.log('ERROR in getHolidayData: ' + error.stack);
        return {
            status: 'error',
            message: `Failed to get holiday data: ${error.message}`,
            data: []
        };
    }
}

/**
 * ดึงข้อมูลพนักงานจาก Google Sheet
 * @returns {object} - อ็อบเจ็กต์พร้อมข้อมูลพนักงาน
 */
function getEmployeeData() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EMPLOYEE_SHEET_NAME);
        if (!sheet || sheet.getLastRow() < 2) {
            return {
                status: 'success',
                data: []
            };
        }

        const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        const employees = values.map(row => {
            const name = row[0];
            const phone = row[1];
            return name ? {
                name: name.trim(),
                phone: (phone || '').toString().trim()
            } : null;
        }).filter(Boolean);

        return {
            status: 'success',
            data: employees
        };
    } catch (error) {
        Logger.log('ERROR in getEmployeeData: ' + error.stack);
        return {
            status: 'error',
            message: `Failed to get employee data: ${error.message}`,
            data: []
        };
    }
}

/**
 * ดึงรายละเอียดของใบงานเฉพาะตามเลขที่ใบสั่งงาน
 * @param {string} docNumber - เลขที่ใบสั่งงานที่ต้องการค้นหา
 * @returns {object} - อ็อบเจ็กต์พร้อมรายละเอียดของใบงาน
 */
function getJobDetailsByDocNumber(docNumber) {
    try {
        if (!docNumber) return {
            status: 'error',
            message: 'Document number is required.'
        };

        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

        const values = sheet.getDataRange().getDisplayValues();
        if (values.length < 2) return {
            status: 'error',
            message: 'No data found.'
        };

        const headers = values[0].map(h => h.trim());
        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        if (docNumberColIndex === -1) return {
            status: 'error',
            message: 'Document number column not found.'
        };

        const row = values.slice(1).find(r => r[docNumberColIndex] === docNumber);
        if (!row) return {
            status: 'error',
            message: `Job with document number ${docNumber} not found.`
        };

        const jobData = headers.reduce((obj, header, index) => {
            if (header) obj[header] = row[index];
            return obj;
        }, {});

        return {
            status: 'success',
            data: jobData
        };
    } catch (error) {
        Logger.log('ERROR in getJobDetailsByDocNumber: ' + error.stack);
        return {
            status: 'error',
            message: `Failed to get job details: ${error.message}`
        };
    }
}


// =================================================================
//                      DATA MODIFICATION FUNCTIONS
// =================================================================

/**
 * เพิ่มใบงานใหม่ลงใน Google Sheet
 * @param {object} data - ข้อมูลใบงานใหม่จากฟอร์ม
 * @returns {object} - สถานะการดำเนินการ
 */
function addNewRepairJob(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const timestamp = new Date();

        const docNumber = generateDocNumber(sheet, timestamp, 'PS');
        const jobFolder = getOrCreateFolder(DriveApp.getFolderById(FOLDER_ID), docNumber);

        const jobPhotoIds = (data.jobPhotos || []).map((photoB64, index) => {
            return saveFileAndGetId(photoB64, `job_${docNumber}_${index + 1}.png`, jobFolder);
        }).filter(id => id).join(',');

        const newRowData = {
            'เลขที่ใบสั่งงาน': docNumber,
            'เลขที่เอกสารอ้างอิง': data.refDocNumber || '',
            'วันที่แจ้ง': Utilities.formatDate(timestamp, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss"),
            'ผู้แจ้ง': data.requesterName || '',
            'ฝ่าย': data.department || '',
            'แผนก': data.division || '',
            'วันที่ให้ดำเนินการ': data.datePerformed ? new Date(data.datePerformed).toISOString() : '',
            'ความต้องการ': data.urgency || 'ปกติ',
            'ประเภทการแจ้ง': data.requestType || '',
            'ชื่อบริษัท': `${data.companyName || ''} ${data.typeofwork || ''}`,
            'โครงการ': data.projectName || '',
            'ผู้ติดต่อ': data.contactPerson || '',
            'เบอร์ติดต่อ': data.contactNumber || '',
            'แผนที่ (Link)': data.mapLink || '',
            'ความพร้อมหน้างาน': data.siteIsReady === null ? '' : (data.siteIsReady ? 'พร้อม' : 'ไม่พร้อม'),
            'รูปภาพหน้างาน': jobPhotoIds,
            'รายละเอียดงาน': data.jobDetails || '',
            'สถานะ': (data.operatorNames && data.operatorNames.length > 0) ? "กำลังดำเนินการ" : "รอดำเนินการ",
            'ผู้ดำเนินการ': (data.operatorNames || []).join(', '),
            'ไฟล์แนบ': '[]' // <-- ADDED: Initialize attachments column
        };

        sheet.appendRow(headers.map(header => newRowData[header] !== undefined ? newRowData[header] : ''));

        return {
            status: 'success',
            message: 'บันทึกข้อมูลเรียบร้อย',
            docNumber: docNumber
        };
    } catch (e) {
        Logger.log('ERROR in addNewRepairJob: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาดในการบันทึก: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * อัปเดตข้อมูลใบงานที่มีอยู่
 * @param {object} data - ข้อมูลใบงานที่ต้องการอัปเดต
 * @returns {object} - สถานะการดำเนินการ
 */
function updateRepairJob(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const docNumber = data.docNumber;
        if (!docNumber) return {
            status: 'error',
            message: 'ไม่พบเลขที่ใบงานสำหรับอัปเดต'
        };

        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        const dataValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const rowIndex = dataValues.findIndex(row => row[docNumberColIndex] == docNumber);
        if (rowIndex === -1) return {
            status: 'error',
            message: `ไม่พบใบงานเลขที่ ${docNumber}`
        };

        const jobFolder = getOrCreateFolder(DriveApp.getFolderById(FOLDER_ID), docNumber);
        const photoData = data.jobPhotos || {
            existing: [],
            new: []
        };
        const existingPhotoIdsOnSheet = (dataValues[rowIndex][headers.indexOf('รูปภาพหน้างาน')] || '').split(',').filter(id => id);
        const photosToKeep = new Set(photoData.existing || []);

        existingPhotoIdsOnSheet.forEach(fileId => {
            if (!photosToKeep.has(fileId)) {
                try {
                    DriveApp.getFileById(fileId).setTrashed(true);
                } catch (e) {
                    Logger.log(`Could not trash file ${fileId}. Error: ${e.message}`);
                }
            }
        });

        const newPhotoIds = (photoData.new || []).map((photoB64, index) => {
            return saveFileAndGetId(photoB64, `job_${docNumber}_update_${Date.now()}_${index + 1}.png`, jobFolder);
        }).filter(id => id);
        const finalPhotoIds = [...photosToKeep, ...newPhotoIds].join(',');

        const currentStatus = dataValues[rowIndex][headers.indexOf('สถานะ')];

        const updatedRowData = {
            'เลขที่เอกสารอ้างอิง': data.refDocNumber,
            'ผู้แจ้ง': data.requesterName,
            'ฝ่าย': data.department,
            'แผนก': data.division,
            'วันที่ให้ดำเนินการ': data.datePerformed ? new Date(data.datePerformed).toISOString() : '',
            'ความต้องการ': data.urgency,
            'ประเภทการแจ้ง': data.requestType,
            'ชื่อบริษัท': data.companyName,
            'โครงการ': data.projectName,
            'ผู้ติดต่อ': data.contactPerson,
            'เบอร์ติดต่อ': data.contactNumber,
            'แผนที่ (Link)': data.mapLink,
            'ความพร้อมหน้างาน': data.siteIsReady === null ? '' : (data.siteIsReady ? 'พร้อม' : 'ไม่พร้อม'),
            'รูปภาพหน้างาน': finalPhotoIds,
            'รายละเอียดงาน': data.jobDetails,
            'ผู้ดำเนินการ': (data.operatorNames || []).join(', ')
        };

        if (currentStatus === 'รอดำเนินการ' && data.operatorNames && data.operatorNames.length > 0) {
            updatedRowData['สถานะ'] = 'กำลังดำเนินการ';
        }

        const rowToUpdate = rowIndex + 2;
        headers.forEach((header, index) => {
            if (updatedRowData[header] !== undefined) {
                sheet.getRange(rowToUpdate, index + 1).setValue(updatedRowData[header]);
            }
        });

        return {
            status: 'success',
            message: 'อัปเดตข้อมูลเรียบร้อย',
            docNumber: docNumber
        };
    } catch (e) {
        Logger.log('ERROR in updateRepairJob: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาดในการอัปเดต: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * มอบหมายงานให้ผู้ดำเนินการ
 * @param {object} data - ข้อมูลการมอบหมายงาน
 * @returns {object} - สถานะการดำเนินการ
 */
function assignJob(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const {
            docNumber,
            operatorNames,
            vehicleLicense
        } = data;

        if (!docNumber || !operatorNames || operatorNames.length === 0) {
            return {
                status: 'error',
                message: 'ข้อมูลไม่ครบถ้วน, กรุณาระบุผู้ดำเนินการอย่างน้อย 1 คน'
            };
        }

        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        const dataValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const rowIndex = dataValues.findIndex(row => row[docNumberColIndex] == docNumber);

        if (rowIndex === -1) {
            return {
                status: 'error',
                message: `ไม่พบใบงานเลขที่ ${docNumber}`
            };
        }

        const rowToUpdate = rowIndex + 2;
        const operatorString = operatorNames.join(', ');

        const fieldsToUpdate = {
            'ผู้ดำเนินการ': operatorString,
            'รถที่ใช้': vehicleLicense || '',
            'สถานะ': 'กำลังดำเนินการ'
        };

        headers.forEach((header, index) => {
            if (fieldsToUpdate[header] !== undefined) {
                sheet.getRange(rowToUpdate, index + 1).setValue(fieldsToUpdate[header]);
            }
        });

        return {
            status: 'success',
            message: 'มอบหมายงานเรียบร้อย',
            docNumber: docNumber
        };
    } catch (e) {
        Logger.log('ERROR in assignJob: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาด: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * อัปเดตประวัติการดำเนินงาน (Action History)
 * @param {object} data - ข้อมูลการดำเนินการ
 * @returns {object} - สถานะการดำเนินการ
 */
function updateJobActionV2(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const docNumber = data.docNumber;
        if (!docNumber) return {
            status: 'error',
            message: 'ไม่พบเลขที่ใบงานสำหรับอัปเดต'
        };

        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        if (docNumberColIndex === -1) return {
            status: 'error',
            message: 'ไม่พบคอลัมน์ "เลขที่ใบสั่งงาน" ใน Sheet'
        };

        const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const rowIndex = allData.findIndex(row => row[docNumberColIndex] == docNumber);
        if (rowIndex === -1) return {
            status: 'error',
            message: `ไม่พบใบงานเลขที่ ${docNumber}`
        };

        const rowToUpdate = rowIndex + 2;
        const historyColIndex = headers.indexOf('ActionHistory');
        let history = [];
        if (historyColIndex !== -1 && allData[rowIndex][historyColIndex]) {
            try {
                history = JSON.parse(allData[rowIndex][historyColIndex]);
            } catch (e) {
                Logger.log(`Could not parse history for ${docNumber}: ${e.message}`);
                history = [];
            }
        }

        const entryDate = data.actionDateTime ? new Date(data.actionDateTime) : new Date();
        const operatorNames = Array.isArray(data.operatorNames) ? data.operatorNames.join(', ') : '';

        const newHistoryEntry = {
            date: entryDate.toISOString(),
            status: data.actionType === 'complete' ? 'ดำเนินการเสร็จสิ้น' : 'กำลังดำเนินการ',
            operator: operatorNames,
            remark: data.actionType === 'complete' ? data.completionNotes : data.remark,
            vehicle: data.vehicleLicense || '',
            departureFromOfficeTime: data.departureFromOfficeTime || '',
            departureFromSiteTime: data.departureFromSiteTime || '',
            arrivalAtCompanyTime: data.arrivalAtCompanyTime || ''
        };

        history.push(newHistoryEntry);

        const fields = {
            'ActionHistory': JSON.stringify(history, null, 2),
            'ผู้ดำเนินการ': operatorNames,
            'รถที่ใช้': data.vehicleLicense || ''
        };

        if (data.actionType === 'incomplete') {
            fields['สถานะ'] = 'กำลังดำเนินการ';
            if (data.nextAppointmentDate) fields['NextAppointment'] = new Date(data.nextAppointmentDate).toISOString();
        } else if (data.actionType === 'complete') {
            fields['สถานะ'] = 'ดำเนินการเสร็จสิ้น';
            fields['NextAppointment'] = '';
            fields['TotalManHours'] = '';
        }

        headers.forEach((header, index) => {
            if (fields[header] !== undefined) {
                sheet.getRange(rowToUpdate, index + 1).setValue(fields[header]);
            }
        });

        return {
            status: 'success',
            message: 'อัปเดตข้อมูลการดำเนินการเรียบร้อย',
            docNumber: docNumber
        };
    } catch (e) {
        Logger.log('ERROR in updateJobActionV2: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาดในการอัปเดต: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * อัปเดตเฉพาะสถานะของใบงาน
 * @param {object} data - ข้อมูลที่ประกอบด้วย docNumber และ status ใหม่
 * @returns {object} - สถานะการดำเนินการ
 */
function updateJobStatus(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const {
            docNumber,
            status: newStatus
        } = data;
        if (!docNumber || !newStatus) return {
            status: 'error',
            message: 'ข้อมูลไม่ครบถ้วนสำหรับอัปเดตสถานะ'
        };

        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        if (docNumberColIndex === -1) return {
            status: 'error',
            message: 'ไม่พบคอลัมน์ "เลขที่ใบสั่งงาน" ใน Sheet'
        };

        const dataValues = sheet.getRange(2, docNumberColIndex + 1, sheet.getLastRow() - 1, 1).getValues();
        const rowIndex = dataValues.findIndex(row => row[0] == docNumber);
        if (rowIndex === -1) return {
            status: 'error',
            message: `ไม่พบใบงานเลขที่ ${docNumber}`
        };

        const statusColIndex = headers.indexOf('สถานะ');
        if (statusColIndex === -1) return {
            status: 'error',
            message: 'ไม่พบคอลัมน์ "สถานะ" ใน Sheet'
        };

        sheet.getRange(rowIndex + 2, statusColIndex + 1).setValue(newStatus);

        return {
            status: 'success',
            message: 'อัปเดตสถานะเรียบร้อย',
            docNumber: docNumber
        };
    } catch (e) {
        Logger.log('ERROR in updateJobStatus: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาดในการอัปเดตสถานะ: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * ลบรายการประวัติการดำเนินงาน
 * @param {string} docNumber - เลขที่ใบงาน
 * @param {number} historyIndex - ลำดับของประวัติที่ต้องการลบ
 * @returns {object} - สถานะการดำเนินการ
 */
function deleteHistoryEntry(docNumber, historyIndex) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());

        if (!docNumber || historyIndex === undefined || historyIndex < 0) {
            return {
                status: 'error',
                message: 'ข้อมูลไม่ครบถ้วนสำหรับการลบประวัติ'
            };
        }

        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        if (docNumberColIndex === -1) return {
            status: 'error',
            message: 'ไม่พบคอลัมน์ "เลขที่ใบสั่งงาน" ใน Sheet'
        };

        const historyColIndex = headers.indexOf('ActionHistory');
        if (historyColIndex === -1) return {
            status: 'error',
            message: 'ไม่พบคอลัมน์ "ActionHistory" ใน Sheet'
        };

        const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const rowIndex = allData.findIndex(row => row[docNumberColIndex] == docNumber);
        if (rowIndex === -1) return {
            status: 'error',
            message: `ไม่พบใบงานเลขที่ ${docNumber}`
        };

        const rowToUpdate = rowIndex + 2;
        let history = [];
        const historyJSON = allData[rowIndex][historyColIndex];

        if (historyJSON) {
            try {
                history = JSON.parse(historyJSON);
            } catch (e) {
                Logger.log(`Could not parse history for ${docNumber} during deletion: ${e.message}`);
                return {
                    status: 'error',
                    message: `ไม่สามารถประมวลผลข้อมูลประวัติได้: ${e.message}`
                };
            }
        }

        if (!Array.isArray(history) || historyIndex >= history.length) {
            return {
                status: 'error',
                message: `ไม่พบประวัติรายการที่ ${historyIndex} ที่จะลบ`
            };
        }

        history.splice(historyIndex, 1);

        sheet.getRange(rowToUpdate, historyColIndex + 1).setValue(JSON.stringify(history, null, 2));

        return {
            status: 'success',
            message: 'ลบประวัติเรียบร้อย'
        };

    } catch (e) {
        Logger.log('ERROR in deleteHistoryEntry: ' + e.stack);
        return {
            status: 'error',
            message: `เกิดข้อผิดพลาดในการลบประวัติ: ${e.message}`
        };
    } finally {
        lock.releaseLock();
    }
}

// =================================================================
//                      FILE UPLOAD FUNCTIONS (NEW)
// =================================================================

/**
 * อัปโหลดไฟล์และเชื่อมโยงกับใบงาน
 * @param {object} fileData - ข้อมูลไฟล์ที่ประกอบด้วย docNumber, fileName, mimeType, base64Data
 * @returns {object} - สถานะการอัปโหลดและข้อมูลไฟล์ที่บันทึก
 */
function uploadFileAndLinkToJob(fileData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const { docNumber, fileName, mimeType, base64Data } = fileData;
        if (!docNumber || !fileName || !mimeType || !base64Data) {
            return { status: 'error', message: 'ข้อมูลสำหรับอัปโหลดไฟล์ไม่ครบถ้วน' };
        }

        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
        
        const docNumberColIndex = headers.indexOf('เลขที่ใบสั่งงาน');
        if (docNumberColIndex === -1) throw new Error('ไม่พบคอลัมน์ "เลขที่ใบสั่งงาน"');

        const attachmentsColName = 'ไฟล์แนบ';
        let attachmentsColIndex = headers.indexOf(attachmentsColName);
        if (attachmentsColIndex === -1) {
            sheet.getRange(1, headers.length + 1).setValue(attachmentsColName);
            attachmentsColIndex = headers.length;
        }

        const dataValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const rowIndex = dataValues.findIndex(row => row[docNumberColIndex] == docNumber);
        if (rowIndex === -1) throw new Error(`ไม่พบใบงานเลขที่ ${docNumber}`);

        const jobFolder = getOrCreateFolder(DriveApp.getFolderById(FOLDER_ID), docNumber);
        const savedFile = saveFileAndGetId(base64Data, fileName, jobFolder, mimeType);
        if (!savedFile.id) throw new Error('ไม่สามารถบันทึกไฟล์ลงใน Drive ได้');

        const rowToUpdate = rowIndex + 2;
        const attachmentsCell = sheet.getRange(rowToUpdate, attachmentsColIndex + 1);
        const currentAttachmentsJSON = attachmentsCell.getValue() || '[]';
        let attachments = [];
        try {
            attachments = JSON.parse(currentAttachmentsJSON);
            if (!Array.isArray(attachments)) attachments = [];
        } catch (e) {
            attachments = [];
        }

        const newAttachment = {
            id: savedFile.id,
            name: fileName,
            type: mimeType,
            size: savedFile.size,
            uploadedAt: new Date().toISOString()
        };

        attachments.push(newAttachment);
        attachmentsCell.setValue(JSON.stringify(attachments));

        return { status: 'success', message: 'อัปโหลดไฟล์สำเร็จ', fileInfo: newAttachment };

    } catch (e) {
        Logger.log('ERROR in uploadFileAndLinkToJob: ' + e.stack);
        return { status: 'error', message: `เกิดข้อผิดพลาดในการอัปโหลดไฟล์: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}


// =================================================================
//                      UTILITY FUNCTIONS
// =================================================================

/**
 * บันทึกไฟล์ (Base64) ลงใน Google Drive และคืนค่า File ID และข้อมูลไฟล์
 * @param {string} base64 - ข้อมูลไฟล์ในรูปแบบ Base64
 * @param {string} filename - ชื่อไฟล์ที่จะบันทึก
 * @param {Folder} folder - โฟลเดอร์ใน Drive ที่จะบันทึกไฟล์
 * @param {string} [forcedMimeType] - Mime type ของไฟล์ (ถ้ามี)
 * @returns {object} - อ็อบเจ็กต์ที่ประกอบด้วย ID, ขนาด และ URL ของไฟล์ หรืออ็อบเจ็กต์ว่างหากล้มเหลว
 */
function saveFileAndGetId(base64, filename, folder, forcedMimeType) {
    if (!base64 || !base64.includes(',')) return {};
    try {
        const parts = base64.split(',');
        const mimeType = forcedMimeType || parts[0].match(/:(.*?);/)[1];
        const bytes = Utilities.base64Decode(parts[1]);
        const blob = Utilities.newBlob(bytes, mimeType, filename);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return {
            id: file.getId(),
            size: file.getSize(),
            url: file.getUrl()
        };
    } catch (e) {
        Logger.log(`Error saving file ${filename}: ${e.stack}`);
        return {};
    }
}

/**
 * สร้างเลขที่ใบสั่งงานที่ไม่ซ้ำกันตามรูปแบบ PS<YY><MM>-<XXXX>
 * @param {Sheet} sheet - ชีตที่ใช้ตรวจสอบข้อมูล
 * @param {Date} timestamp - วันที่ปัจจุบัน
 * @param {string} prefixCode - รหัสคำนำหน้า (เช่น PS)
 * @returns {string} - เลขที่ใบสั่งงานใหม่
 */
function generateDocNumber(sheet, timestamp, prefixCode) {
    const year = String(timestamp.getFullYear() + 543).slice(-2);
    const month = ('0' + (timestamp.getMonth() + 1)).slice(-2);
    const prefix = `${prefixCode}${year}${month}-`;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const docNumberColIndex = headers.indexOf("เลขที่ใบสั่งงาน");
    if (docNumberColIndex === -1) return `${prefix}0001`;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return `${prefix}0001`;

    const docNumbers = sheet.getRange(2, docNumberColIndex + 1, lastRow - 1, 1)
        .getValues()
        .flat()
        .filter(id => typeof id === 'string' && id.startsWith(prefix))
        .map(id => parseInt(id.slice(prefix.length), 10))
        .filter(num => !isNaN(num));

    const nextNumber = docNumbers.length > 0 ? Math.max(...docNumbers) + 1 : 1;

    return `${prefix}${String(nextNumber).padStart(4, '0')}`;
}

/**
 * ค้นหาหรือสร้างโฟลเดอร์ใน Google Drive
 * @param {Folder} parentFolder - โฟลเดอร์แม่
 * @param {string} folderName - ชื่อโฟลเดอร์ที่ต้องการ
 * @returns {Folder} - โฟลเดอร์ที่พบหรือสร้างใหม่
 */
function getOrCreateFolder(parentFolder, folderName) {
    const folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}
