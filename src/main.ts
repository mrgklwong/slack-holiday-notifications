import ExcelJS from 'exceljs';
import axios from 'axios';
import { format, isSameDay, isWithinInterval, parse } from 'date-fns';

const EXCEL_FILE_PATH = process.argv.includes('--excelPath')
    ? process.argv[process.argv.indexOf('--excelPath') + 1]
    : 'leaves.xlsx';

interface LeaveEntry {
    Employee: string;
    StartDate: Date;
    FinishDate: Date;
    LeaveType: string;
    Status: string;
}

const SLACK_WEBHOOK_URL = process.env.SLACK_WEBHOOK_URL!;

// Custom parser focusing on date portion only
function parseJsDate(dateStr: string): Date {
    const dateParts = dateStr.split(' ').slice(1, 4); // Extract "MMM", "dd", "yyyy"
    return parse(dateParts.join(' '), "MMM dd yyyy", new Date());
}

async function readExcel(): Promise<LeaveEntry[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);

    const sheet = workbook.getWorksheet('VOGSY Data');
    if (!sheet) throw new Error('VOGSY Data sheet not found');

    const entries: LeaveEntry[] = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return;

        const rawStart = row.getCell(2).text;
        const rawEnd = row.getCell(3).text;

        entries.push({
            Employee: row.getCell(1).text,
            StartDate: parseJsDate(rawStart),
            FinishDate: parseJsDate(rawEnd),
            LeaveType: row.getCell(5).text,
            Status: row.getCell(8).text.toLowerCase()
        });
    });

    return entries;
}

async function main() {
    const entries = await readExcel();

    // Format filtered entries for logging
    const filteredEntries = entries.filter(e => e.Employee === 'Test');
    console.log('Test Employee Entries:');
    filteredEntries.forEach(entry => {
        console.log({
            Employee: entry.Employee,
            StartDate: format(entry.StartDate, 'yyyy-MM-dd'),
            FinishDate: format(entry.FinishDate, 'yyyy-MM-dd'),
            LeaveType: entry.LeaveType,
            Status: entry.Status
        });
    });

    const today = new Date();
    console.log(`Today's Date: ${format(today, 'yyyy-MM-dd')}`);

    const relevantLeaves = entries.filter(entry => {
        const isValidLeave = ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType);
        const isActiveStatus = ['approved', 'submitted'].includes(entry.Status);
        return isValidLeave &&
            isActiveStatus &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, { start: entry.StartDate, end: entry.FinishDate }) ||
                isSameDay(today, entry.FinishDate));
    });

    let slackMessage = `:palm_tree: *Daily Leave Report (${format(today, 'yyyy-MM-dd')})* :palm_tree:\n`;

    if (relevantLeaves.length === 0) {
        slackMessage += 'No one is slacking today!';
    } else {
        relevantLeaves.forEach(entry => {
            slackMessage += `• *${entry.Employee}* - ${entry.LeaveType} ` +
                `(${format(entry.StartDate, 'yyyy-MM-dd')} - ` +
                `${format(entry.FinishDate, 'yyyy-MM-dd')})\n`;
        });
    }

    slackMessage += 'Stay sunny! :sun_with_face:';

    console.log('Sending to Slack:', slackMessage);
    await axios.post(SLACK_WEBHOOK_URL, { text: slackMessage, channel: '#testing' });
}

main().catch(console.error);
