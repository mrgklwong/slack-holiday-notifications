import ExcelJS from 'exceljs';
import axios from 'axios';
import { format, isSameDay, isWithinInterval, parse } from 'date-fns';

// Configuration for department-to-channel mapping (update as needed)
const DEPARTMENT_CHANNEL_MAP = {
    UK: '#garytesting',
    India: '#garytesting',
    OZ: '#garytesting'
};

// Common validation for channel mapping
const validateChannelMapping = (): void => {
    const missingChannels = Object.values(DEPARTMENT_CHANNEL_MAP).filter(channel => channel === '');
    if (missingChannels.length > 0) {
        throw new Error('Incomplete channel mapping configuration in DEPARTMENT_CHANNEL_MAP');
    }
};

const EXCEL_FILE_PATH = process.argv.includes('--excelPath')
    ? process.argv[process.argv.indexOf('--excelPath') + 1]
    : 'leaves.xlsx';

type SlackPayload = {
    text: string;
    channel?: string;
};

interface LeaveEntry {
    Employee: string;
    StartDate: Date;
    FinishDate: Date;
    LeaveType: string;
    Status: string;
    Department: string;
}

const SLACK_WEBHOOK_URL = process.env.SLACK_WEBHOOK_URL!;

// Custom parser with better error handling
function parseJsDate(dateStr: string): Date {
    const defaultDate = new Date();
    let date = parse(dateStr, "MMM dd yyyy", defaultDate);

    if (isNaN(date.getTime())) {
        // Try alternative date formats if needed
        date = parse(dateStr, "M/d/yyyy", new Date());
        if (isNaN(date.getTime())) {
            throw new Error(`Invalid date format: ${dateStr}`);
        }
    }

    return date;
}

async function readExcel(): Promise<LeaveEntry[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);

    const sheet = workbook.getWorksheet('VOGSY Data');
    if (!sheet) throw new Error('VOGSY Data sheet not found');

    const headers = sheet.getRow(1).values as any[];
    const employeeCol = headers.findIndex(val => val === 'Employee') + 1;
    const startDateCol = headers.findIndex(val => val === 'Start date') + 1;
    const finishDateCol = headers.findIndex(val => val === 'Finish date') + 1;
    const leaveTypeCol = headers.findIndex(val => val === 'Description leave budget') + 1;
    const statusCol = headers.findIndex(val => val === 'Status') + 1;
    const companyDepCol = headers.findIndex(val => val === 'Company / department') + 1;

    // Validate all required headers exist
    if ([employeeCol, startDateCol, finishDateCol, leaveTypeCol, statusCol, companyDepCol].some(i => i === 0)) {
        throw new Error('Could not find expected columns in Excel sheet');
    }

    const validateCellValue = (value: any): string => {
        if (!value) return '';
        return typeof value === 'string' ? value.trim() : String(value);
    };

    return sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return;

        // Extract and normalize company/department value
        const rawdep = validateCellValue(row.getCell(companyDepCol).value);
        const department = rawdep.split('/').pop()?.trim().toLowerCase() || '';

        // Extract and parse dates
        const startCell = row.getCell(startDateCol);
        const startDateStr = typeof startCell.text === 'string' ? startCell.text : '';
        const startDateVal = startCell.value instanceof Date
            ? startCell.value
            : parseJsDate(startDateStr);

        const finishCell = row.getCell(finishDateCol);
        const finishDateStr = typeof finishCell.text === 'string' ? finishCell.text : '';
        const finishDateVal = finishCell.value instanceof Date
            ? finishCell.value
            : parseJsDate(finishDateStr);

        return {
            Employee: validateCellValue(row.getCell(employeeCol).value),
            StartDate: startDateVal,
            FinishDate: finishDateVal,
            LeaveType: validateCellValue(row.getCell(leaveTypeCol).value).toLowerCase(),
            Status: validateCellValue(row.getCell(statusCol).value).toLowerCase(),
            Department: department
        };
    });
}

async function getActiveLeaves(entries: LeaveEntry[]): Promise<LeaveEntry[]> {
    const today = new Date();
    return entries.filter(entry => {
        return ['Annual Leave', 'Sick Leave'].includes(entry.LeaveType) &&
            ['approved', 'submitted'].includes(entry.Status) &&
            (isSameDay(today, entry.StartDate) ||
                isWithinInterval(today, {
                    start: entry.StartDate,
                    end: entry.FinishDate
                }));
    });
}

function groupLeavesByChannel(entries: LeaveEntry[]): Record<string, LeaveEntry[]> {
    validateChannelMapping();
    const today = format(new Date(), 'dd-MM-yyyy');

    return Object.entries(DEPARTMENT_CHANNEL_MAP).reduce((groups, [department, channel]) => {
        const filtered = entries.filter(entry =>
            entry.Department === department.toLowerCase()
        );

        if (filtered.length > 0) {
            groups[channel] = filtered;
        }
        return groups;
    }, {} as Record<string, LeaveEntry[]>);
}

async function sendSlackMessage(channel: string, leaves: LeaveEntry[]): Promise<void> {
    const formattedLeaves = leaves.map(entry =>
        `• *${entry.Employee}* - ${entry.LeaveType} ` +
        `(${format(entry.StartDate, 'dd-MM-yyyy')} - ` +
        `${format(entry.FinishDate, 'dd-MM-yyyy')})`
    );

    let message = `:palm_tree: *Daily Leave Report (${format(new Date(), 'dd-MM-yyyy')})* :palm_tree:\n`;

    if (leaves.length === 0) {
        message += 'No one is slacking today!';
    } else {
        message += `There are **${leaves.length}** leaves today in ${channel}:\n` + formattedLeaves.join('\n');
    }

    message += '\nStay sunny! :sun_with_face:';

    try {
        const payload: SlackPayload = {
            text: message,
            channel: channel
        };

        console.log(`Sending to Slack channel ${channel}:`, payload.text);
        await axios.post(SLACK_WEBHOOK_URL, payload);
    } catch (error) {
        console.error(`Failed to send to Slack channel ${channel}:`, error.message);
    }
}

async function main() {
    // Read and process leave data from Excel file
    const excelEntries = await readExcel();
    const relevantLeaves = await getActiveLeaves(excelEntries);

    // Group leaves by their department to targeted channels
    const channelGroupedLeaves = groupLeavesByChannel(relevantLeaves);

    for (const [channel, leaves] of Object.entries(channelGroupedLeaves)) {
        await sendSlackMessage(channel, leaves);
    }
}

// Ensure environment variables are set and start the process
if (!SLACK_WEBHOOK_URL) {
    console.error('SLACK_WEBHOOK_URL environment variable is missing');
    process.exit(1);
}

main().catch(console.error);
