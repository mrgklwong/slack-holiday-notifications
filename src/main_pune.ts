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
    Office: string;
}

const SLACK_WEBHOOK_URL = process.env.SLACK_WEBHOOK_URL_IN!;
// Custom parser focusing on date portion only
function parseJsDate(dateStr: string): Date {
    const dateParts = dateStr.split(' ').slice(1, 4); // ["May", "23", "2025"]
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

        const rawStart = row.getCell(3).text;
        const rawEnd = row.getCell(4).text;

        entries.push({
            Employee: row.getCell(1).text,
            StartDate: parseJsDate(rawStart),
            FinishDate: parseJsDate(rawEnd),
            LeaveType: row.getCell(6).text,
            Status: row.getCell(9).text.toLowerCase(),
            Office: row.getCell(2).text.toLowerCase()
        });
    });

    return entries;
}

async function main() {
    const entries = await readExcel();

    // Format filtered entries for logging
    const filteredEntries = entries.filter(e => e.Employee === 'Gary Wong');
    console.log('Test Employee Entries:');
    filteredEntries.forEach(entry => {
        console.log({
            Employee: entry.Employee,
            StartDate: format(entry.StartDate, 'dd/MM/yyyy'),
            FinishDate: format(entry.FinishDate, 'dd/MM/yyyy'),
            LeaveType: entry.LeaveType,
            Status: entry.Status,
            Office: entry.Office
        });
    });

    const today = new Date();
    console.log(`Today's Date: ${today}`);

    const relevantLeaves = entries.filter(entry => {
        const isValidLeave = ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType);
        const isActiveStatus = ['approved', 'submitted'].includes(entry.Status);
        return isActiveStatus &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, { start: entry.StartDate, end: entry.FinishDate }) ||
                isSameDay(today, entry.FinishDate));
    });

    const UKLeaves = entries.filter(entry => {
        const isValidLeave = ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType);
        const isActiveStatus = ['approved', 'submitted'].includes(entry.Status);
        const isUK = ['clearroute uk ltd'].includes((entry.Office));
        return isActiveStatus &&
                isUK &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, { start: entry.StartDate, end: entry.FinishDate }) ||
                isSameDay(today, entry.FinishDate));
    });

    const OZLeaves = entries.filter(entry => {
        const isValidLeave = ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType);
        const isActiveStatus = ['approved', 'submitted'].includes(entry.Status);
        const isUK = ['clearroute australia pty ltd'].includes((entry.Office));
        return isActiveStatus &&
            isUK &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, { start: entry.StartDate, end: entry.FinishDate }) ||
                isSameDay(today, entry.FinishDate));
    });

    const indiaLeaves = entries.filter(entry => {
        const isValidLeave = ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType);
        const isActiveStatus = ['approved', 'submitted'].includes(entry.Status);
        const isUK = ['clearroute india'].includes((entry.Office));
        return isActiveStatus &&
            isUK &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, { start: entry.StartDate, end: entry.FinishDate }) ||
                isSameDay(today, entry.FinishDate));
    });

    let slackMessage = `:palm_tree: *Daily Leave Report (${format(today, 'dd/MM/yyyy')})* :palm_tree:\n`;

    if (indiaLeaves.length === 0) {
        slackMessage += 'No one is slacking today!';
    } else {
        indiaLeaves.forEach(entry => {
            slackMessage += `• *${entry.Employee}* - ${entry.LeaveType} ` +
                `(${format(entry.StartDate, 'dd/MM/yyyy')} - ` +
                `${format(entry.FinishDate, 'dd/MM/yyyy')})\n`;
        });
    }
    slackMessage += 'Stay sunny! :sun_with_face:';
    console.log('Sending to Slack IN:', slackMessage);
    await axios.post(SLACK_WEBHOOK_URL, { text: slackMessage, channel: '#testing' });

}

main().catch(console.error);
