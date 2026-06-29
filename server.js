const express = require('express');
const session = require('cookie-session');
const bodyParser = require('body-parser');
const path = require('path');
const exceljs = require('exceljs');
const docx = require('docx');
const db = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize Database (moved to startServer at the end of the file)

// Middlewares
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

app.use(session({
    name: 'session',
    keys: ['security-permits-key-2026-secret'],
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
}));

// Middleware to ensure database is loaded (critical for serverless like Vercel)
app.use(async (req, res, next) => {
    try {
        await db.initDb();
        next();
    } catch (err) {
        console.error('Failed to initialize database on request:', err);
        res.status(500).send('خطأ في الاتصال بقاعدة البيانات');
    }
});

// Route protection rules for static files
app.get('/', (req, res) => {
    return res.redirect('/index.html');
});

app.get('/index.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/dashboard.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

app.get('/login.html', (req, res) => {
    return res.redirect('/dashboard.html');
});

// Serve public static assets
app.use(express.static(path.join(__dirname, 'public')));

// Authentication middleware for admin API
const requireAdmin = (req, res, next) => {
    next();
};

// API Endpoints

// Admin Login
app.post('/api/login', (req, res) => {
    const { passcode } = req.body;
    if (!passcode) {
        return res.status(400).json({ error: 'يرجى إدخال رمز الدخول للمسؤول' });
    }

    const adminUser = db.getUsers().find(u => u.role === 'admin');
    const validPasscode = adminUser ? adminUser.passcode : 'admin123';

    if (passcode === validPasscode) {
        req.session.userId = adminUser ? adminUser.id : 'admin_default';
        req.session.userName = adminUser ? adminUser.name : 'مدير النظام';
        req.session.role = 'admin';
        
        return res.json({ 
            success: true, 
            role: 'admin', 
            name: adminUser ? adminUser.name : 'مدير النظام',
            redirect: '/dashboard.html'
        });
    } else {
        return res.status(401).json({ error: 'رمز الدخول غير صحيح' });
    }
});

// Logout
app.post('/api/logout', (req, res) => {
    req.session = null;
    res.json({ success: true });
});

// Check Session
app.get('/api/session', (req, res) => {
    if (req.session.userId) {
        res.json({
            loggedIn: true,
            userId: req.session.userId,
            name: req.session.userName,
            role: req.session.role
        });
    } else {
        res.json({ loggedIn: false });
    }
});

// Change Admin Passcode (Admin only)
app.post('/api/admin/change-passcode', requireAdmin, (req, res) => {
    const { currentPasscode, newPasscode } = req.body;
    if (!currentPasscode || !newPasscode) {
        return res.status(400).json({ error: 'الرجاء إدخال رمز الدخول الحالي والجديد' });
    }

    try {
        const adminUser = db.getUsers().find(u => u.role === 'admin');
        const validPasscode = adminUser ? adminUser.passcode : 'admin123';

        if (currentPasscode !== validPasscode) {
            return res.status(400).json({ error: 'رمز الدخول الحالي غير صحيح' });
        }

        db.updateAdminPasscode(newPasscode);
        res.json({ success: true, message: 'تم تغيير رمز الدخول بنجاح' });
    } catch (e) {
        console.error('Error changing admin passcode:', e);
        res.status(500).json({ error: 'حدث خطأ أثناء تغيير رمز الدخول' });
    }
});

// Form Submissions

// Post Form (Public - No Auth required)
app.post('/api/submit', (req, res) => {
    const { userName, governorate, month, rows } = req.body;
    if (!userName || !governorate || !month || !rows || !Array.isArray(rows) || rows.length === 0) {
        return res.status(400).json({ error: 'بيانات الاستمارة غير مكتملة أو فارغة' });
    }

    try {
        const submission = db.addSubmission(
            'public_employee', 
            userName.trim(), 
            governorate, 
            month, 
            rows
        );
        res.json({ success: true, submission });
    } catch (e) {
        res.status(500).json({ error: 'حدث خطأ أثناء حفظ الاستمارة' });
    }
});

// Get Submissions (Admin only)
app.get('/api/submissions', requireAdmin, (req, res) => {
    try {
        let submissions = db.getSubmissions();
        const { user, month, governorate } = req.query;

        if (user) {
            submissions = submissions.filter(s => s.userName === user);
        }
        if (month) {
            submissions = submissions.filter(s => s.month === month);
        }
        if (governorate) {
            submissions = submissions.filter(s => s.governorate === governorate);
        }

        // Sort by date descending
        submissions.sort((a, b) => new Date(b.submissionDate) - new Date(a.submissionDate));

        res.json(submissions);
    } catch (e) {
        res.status(500).json({ error: 'خطأ في جلب الاستمارات' });
    }
});

// Get My Submissions (Public - used by employees to view their history)
app.get('/api/my-submissions', (req, res) => {
    const { name } = req.query;
    if (!name || name.trim() === '') {
        return res.json([]);
    }

    try {
        const queryName = name.trim().toLowerCase();
        const submissions = db.getSubmissions().filter(s => 
            s.userName && s.userName.toLowerCase() === queryName
        );

        // Sort by date descending
        submissions.sort((a, b) => new Date(b.submissionDate) - new Date(a.submissionDate));

        res.json(submissions);
    } catch (e) {
        res.status(500).json({ error: 'خطأ في جلب استمارات الموظف' });
    }
});

// Delete Submission (Admin only)
app.delete('/api/submissions/:id', requireAdmin, (req, res) => {
    try {
        const success = db.deleteSubmission(req.params.id);
        if (success) {
            res.json({ success: true });
        } else {
            res.status(404).json({ error: 'الاستمارة غير موجودة' });
        }
    } catch (e) {
        res.status(500).json({ error: 'خطأ في حذف الاستمارة' });
    }
});

// Filter Helper for Exports
function filterSubmissions(query) {
    let submissions = db.getSubmissions();
    const { user, month, governorate } = query;

    if (user) {
        submissions = submissions.filter(s => s.userName === user);
    }
    if (month) {
        submissions = submissions.filter(s => s.month === month);
    }
    if (governorate) {
        submissions = submissions.filter(s => s.governorate === governorate);
    }

    submissions.sort((a, b) => new Date(a.submissionDate) - new Date(b.submissionDate));
    return submissions;
}

// Export to Excel (Grouped by Governorate, with styled headers mirroring index.html layout)
// Export to Excel (Grouped and deduplicated, with styled headers mirroring index.html layout)
app.get('/api/export/excel', requireAdmin, async (req, res) => {
    try {
        const submissions = filterSubmissions(req.query);

        if (submissions.length === 0) {
            return res.status(404).send('لا توجد بيانات لتصديرها');
        }

        const workbook = new exceljs.Workbook();

        // Group by governorate and month, and deduplicate rows
        const govGroups = {};
        submissions.forEach(sub => {
            const gov = sub.governorate || 'غير محدد';
            const m = sub.month || 'غير محدد';
            if (!govGroups[gov]) {
                govGroups[gov] = {};
            }
            if (!govGroups[gov][m]) {
                govGroups[gov][m] = [];
            }
            if (sub.rows && Array.isArray(sub.rows)) {
                govGroups[gov][m] = govGroups[gov][m].concat(sub.rows);
            }
        });

        // Deduplicate rows function
        function getUniqueRows(rows) {
            const seen = new Set();
            return rows.filter(row => {
                const key = [
                    (row.tasks || '').trim(),
                    (row.employees || '').toString().trim(),
                    (row.planned || '').toString().trim(),
                    (row.completion || '').trim(),
                    (row.section || '').trim(),
                    (row.department || '').trim(),
                    (row.beneficiary || '').trim(),
                    (row.supporting || '').trim(),
                    (row.date || '').trim()
                ].join('|||').toLowerCase();
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        // Generate sheets grouped by governorate
        for (const govName in govGroups) {
            const safeSheetName = govName.replace(/[:\\/?*\[\]]/g, '').substring(0, 30) || 'صفحة';
            const worksheet = workbook.addWorksheet(safeSheetName, {
                views: [{ rtl: true }]
            });

            worksheet.columns = [
                { key: 'seq', width: 6 },
                { key: 'tasks', width: 45 },
                { key: 'employees_count', width: 22 },
                { key: 'planned', width: 22 },
                { key: 'completion', width: 22 },
                { key: 'section', width: 18 },
                { key: 'dept', width: 18 },
                { key: 'beneficiary', width: 20 },
                { key: 'supporting', width: 20 },
                { key: 'exec_date', width: 15 }
            ];

            let currentRow = 1;
            let monthIdx = 0;

            for (const monthName in govGroups[govName]) {
                const uniqueRows = getUniqueRows(govGroups[govName][monthName]);
                if (uniqueRows.length === 0) continue;

                if (monthIdx > 0) {
                    currentRow += 3; // spacing between tables
                }

                // Row 1: Banner Header - "قسم التصاريح الأمنية"
                worksheet.mergeCells(currentRow, 1, currentRow, 10);
                const titleCell1 = worksheet.getCell(currentRow, 1);
                titleCell1.value = 'قسم التصاريح الأمنية';
                titleCell1.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFF' } };
                titleCell1.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '1E3C72' }
                };
                titleCell1.alignment = { horizontal: 'center', vertical: 'middle' };
                worksheet.getRow(currentRow).height = 25;

                currentRow++;

                // Row 2: Banner Subtitle 1 - "شعبة المتابعة"
                worksheet.mergeCells(currentRow, 1, currentRow, 10);
                const titleCellSub1 = worksheet.getCell(currentRow, 1);
                titleCellSub1.value = 'شعبة المتابعة';
                titleCellSub1.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
                titleCellSub1.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '1E3C72' }
                };
                titleCellSub1.alignment = { horizontal: 'center', vertical: 'middle' };
                worksheet.getRow(currentRow).height = 20;

                currentRow++;

                // Row 3: Banner Subtitle 2 - "استمارة الأنشطة الشهرية"
                worksheet.mergeCells(currentRow, 1, currentRow, 10);
                const titleCell2 = worksheet.getCell(currentRow, 1);
                titleCell2.value = 'استمارة الأنشطة الشهرية';
                titleCell2.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
                titleCell2.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '1E3C72' }
                };
                titleCell2.alignment = { horizontal: 'center', vertical: 'middle' };
                worksheet.getRow(currentRow).height = 20;

                currentRow++;

                // Row 3: Banner Location & Date - "محافظة (...) - لشهر (...)"
                worksheet.mergeCells(currentRow, 1, currentRow, 10);
                const titleCell3 = worksheet.getCell(currentRow, 1);
                titleCell3.value = `محافظة (${govName}) - لشهر (${monthName})`;
                titleCell3.font = { name: 'Arial', size: 11, color: { argb: 'FFFFFF' } };
                titleCell3.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '1E3C72' }
                };
                titleCell3.alignment = { horizontal: 'center', vertical: 'middle' };
                worksheet.getRow(currentRow).height = 20;

                currentRow++;

                // Row 4: Column headers
                const headers = [
                    'ت',
                    'المهام والواجبات',
                    'عدد الموظفين القائمين بالخدمة',
                    'عدد الأعمال ضمن الخطة',
                    'نسبة الإنجاز أو عدد المنجز',
                    'القسم',
                    'الشعبة',
                    'الجهة المستفيدة',
                    'الجهة الساندة',
                    'تاريخ التنفيذ'
                ];

                const headerRow = worksheet.getRow(currentRow);
                headerRow.height = 30;
                headers.forEach((hText, hIdx) => {
                    const cell = headerRow.getCell(hIdx + 1);
                    cell.value = hText;
                    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: '2A5298' }
                    };
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'CCCCCC' } },
                        left: { style: 'thin', color: { argb: 'CCCCCC' } },
                        bottom: { style: 'medium', color: { argb: '1E3C72' } },
                        right: { style: 'thin', color: { argb: 'CCCCCC' } }
                    };
                });

                currentRow++;

                // Add data rows
                uniqueRows.forEach((row, rIdx) => {
                    const dataRow = worksheet.getRow(currentRow);
                    dataRow.height = 25;

                    const rowValues = [
                        rIdx + 1,
                        row.tasks,
                        Number(row.employees) || row.employees || '',
                        Number(row.planned) || row.planned || '',
                        row.completion,
                        row.section,
                        row.department,
                        row.beneficiary,
                        row.supporting,
                        row.date
                    ];

                    rowValues.forEach((val, valIdx) => {
                        const cell = dataRow.getCell(valIdx + 1);
                        cell.value = val;
                        cell.font = { name: 'Arial', size: 10 };
                        cell.alignment = { 
                            horizontal: valIdx === 1 ? 'right' : 'center', 
                            vertical: 'middle', 
                            wrapText: true 
                        };
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'E5E5E5' } },
                            left: { style: 'thin', color: { argb: 'E5E5E5' } },
                            bottom: { style: 'thin', color: { argb: 'E5E5E5' } },
                            right: { style: 'thin', color: { argb: 'E5E5E5' } }
                        };
                    });

                    currentRow++;
                });

                monthIdx++;
            }
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Report_Cumulative_${Date.now()}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (e) {
        console.error('Error generating Excel:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Excel: ' + e.message);
    }
});

// Export to Word (Grouped and deduplicated, landscape pages)
app.get('/api/export/word', requireAdmin, async (req, res) => {
    try {
        const submissions = filterSubmissions(req.query);

        if (submissions.length === 0) {
            return res.status(404).send('لا توجد بيانات لتصديرها');
        }

        const { Table, TableRow, TableCell, Document, Paragraph, TextRun, AlignmentType, WidthType, Packer } = docx;

        const docChildren = [];

        // Group by governorate and month
        const govGroups = {};
        submissions.forEach(sub => {
            const gov = sub.governorate || 'غير محدد';
            const m = sub.month || 'غير محدد';
            if (!govGroups[gov]) govGroups[gov] = {};
            if (!govGroups[gov][m]) govGroups[gov][m] = [];
            if (sub.rows && Array.isArray(sub.rows)) {
                govGroups[gov][m] = govGroups[gov][m].concat(sub.rows);
            }
        });

        // Deduplicate rows function
        function getUniqueRows(rows) {
            const seen = new Set();
            return rows.filter(row => {
                const key = [
                    (row.tasks || '').trim(),
                    (row.employees || '').toString().trim(),
                    (row.planned || '').toString().trim(),
                    (row.completion || '').trim(),
                    (row.section || '').trim(),
                    (row.department || '').trim(),
                    (row.beneficiary || '').trim(),
                    (row.supporting || '').trim(),
                    (row.date || '').trim()
                ].join('|||').toLowerCase();
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        let isFirst = true;

        for (const govName in govGroups) {
            for (const monthName in govGroups[govName]) {
                const uniqueRows = getUniqueRows(govGroups[govName][monthName]);
                if (uniqueRows.length === 0) continue;

                if (!isFirst) {
                    docChildren.push(new Paragraph({ text: '', spacing: { before: 200, after: 200 } }));
                }
                isFirst = false;

                // Table header block matching the HTML banner styling
                docChildren.push(
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100, after: 50 },
                        children: [
                            new TextRun({ text: 'قسم التصاريح الأمنية', bold: true, color: '1E3C72', size: 24 })
                        ]
                    })
                );

                docChildren.push(
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 50 },
                        children: [
                            new TextRun({ text: 'شعبة المتابعة', bold: true, color: '1E3C72', size: 18 })
                        ]
                    })
                );

                docChildren.push(
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 50 },
                        children: [
                            new TextRun({ text: 'استمارة الأنشطة الشهرية', bold: true, color: '1E3C72', size: 20 })
                        ]
                    })
                );

                docChildren.push(
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 200 },
                        children: [
                            new TextRun({ text: `محافظة (${govName}) - لشهر (${monthName})`, bold: true, color: '2A5298', size: 16 })
                        ]
                    })
                );

                // Create Table
                const tableRows = [];

                // 10 Columns header
                const headerRow = new TableRow({
                    children: [
                        new TableCell({ width: { size: 400, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'ت', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'المهام والواجبات', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الموظفين', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الأعمال', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'نسبة الإنجاز', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'القسم', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الشعبة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة المستفيدة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة الساندة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'تاريخ التنفيذ', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] })
                    ]
                });
                tableRows.push(headerRow);

                // Add data rows
                uniqueRows.forEach((r, rIdx) => {
                    const cells = [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(rIdx + 1), size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.tasks || '', size: 15 })], alignment: AlignmentType.RIGHT })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.employees || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.planned || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.completion || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.section || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.department || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.beneficiary || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.supporting || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.date || '', size: 15 })], alignment: AlignmentType.CENTER })] })
                    ];
                    tableRows.push(new TableRow({ children: cells }));
                });

                const wordTable = new Table({
                    rows: tableRows,
                    width: { size: 100, type: WidthType.PERCENTAGE }
                });

                docChildren.push(wordTable);
            }
        }

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 800, bottom: 800, left: 800, right: 800 },
                        size: { orientation: 'landscape' }
                    }
                },
                children: docChildren
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename=Report_Cumulative_${Date.now()}.docx`);
        res.send(buffer);
    } catch (e) {
        console.error('Error generating Word:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Word: ' + e.message);
    }
});

// Export individual submission as Word (Public)
app.get('/api/export/employee/word', async (req, res) => {
    try {
        const { id, name } = req.query;
        if (!id || !name) {
            return res.status(400).send('طلب غير صالح');
        }

        const submission = db.getSubmissions().find(s => s.id === id && s.userName.toLowerCase() === name.trim().toLowerCase());
        if (!submission) {
            return res.status(404).send('الاستمارة غير موجودة');
        }

        const { Table, TableRow, TableCell, Document, Paragraph, TextRun, AlignmentType, WidthType, Packer } = docx;
        const docChildren = [];

        // Header block
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 50 }, children: [new TextRun({ text: 'قسم التصاريح الأمنية', bold: true, color: '1E3C72', size: 24 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 50 }, children: [new TextRun({ text: 'شعبة المتابعة', bold: true, color: '1E3C72', size: 18 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 50 }, children: [new TextRun({ text: 'استمارة الأنشطة الشهرية', bold: true, color: '1E3C72', size: 20 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 }, children: [new TextRun({ text: `محافظة (${submission.governorate}) - لشهر (${submission.month})`, bold: true, color: '2A5298', size: 16 })] }));

        const tableRows = [];
        const headerRow = new TableRow({
            children: [
                new TableCell({ width: { size: 400, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'ت', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'المهام والواجبات', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الموظفين', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الأعمال', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'نسبة الإنجاز', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'القسم', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الشعبة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة المستفيدة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة الساندة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'تاريخ التنفيذ', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] })
            ]
        });
        tableRows.push(headerRow);

        submission.rows.forEach((r, rIdx) => {
            const cells = [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(rIdx + 1), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.tasks || '', size: 15 })], alignment: AlignmentType.RIGHT })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.employees || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.planned || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.completion || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.section || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.department || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.beneficiary || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.supporting || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.date || '', size: 15 })], alignment: AlignmentType.CENTER })] })
            ];
            tableRows.push(new TableRow({ children: cells }));
        });

        const wordTable = new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
        docChildren.push(wordTable);

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 800, bottom: 800, left: 800, right: 800 }, size: { orientation: 'landscape' } } },
                children: docChildren
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent(`Report_${submission.governorate}_${submission.month}.docx`));
        res.send(buffer);
    } catch (e) {
        console.error('Error generating Word receipt:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Word: ' + e.message);
    }
});

// Export individual submission as Excel (Public)
app.get('/api/export/employee/excel', async (req, res) => {
    try {
        const { id, name } = req.query;
        if (!id || !name) {
            return res.status(400).send('طلب غير صالح');
        }

        const submission = db.getSubmissions().find(s => s.id === id && s.userName.toLowerCase() === name.trim().toLowerCase());
        if (!submission) {
            return res.status(404).send('الاستمارة غير موجودة');
        }

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet(submission.governorate, { views: [{ rtl: true }] });

        worksheet.columns = [
            { key: 'seq', width: 6 },
            { key: 'tasks', width: 45 },
            { key: 'employees_count', width: 22 },
            { key: 'planned', width: 22 },
            { key: 'completion', width: 22 },
            { key: 'section', width: 18 },
            { key: 'dept', width: 18 },
            { key: 'beneficiary', width: 20 },
            { key: 'supporting', width: 20 },
            { key: 'exec_date', width: 15 }
        ];

        worksheet.mergeCells(1, 1, 1, 10);
        const titleCell1 = worksheet.getCell(1, 1);
        titleCell1.value = 'قسم التصاريح الأمنية';
        titleCell1.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFF' } };
        titleCell1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell1.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 25;

        worksheet.mergeCells(2, 1, 2, 10);
        const titleCellSub1 = worksheet.getCell(2, 1);
        titleCellSub1.value = 'شعبة المتابعة';
        titleCellSub1.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        titleCellSub1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCellSub1.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(2).height = 20;

        worksheet.mergeCells(3, 1, 3, 10);
        const titleCell2 = worksheet.getCell(3, 1);
        titleCell2.value = 'استمارة الأنشطة الشهرية';
        titleCell2.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        titleCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell2.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(3).height = 20;

        worksheet.mergeCells(4, 1, 4, 10);
        const titleCell3 = worksheet.getCell(4, 1);
        titleCell3.value = `محافظة (${submission.governorate}) - لشهر (${submission.month})`;
        titleCell3.font = { name: 'Arial', size: 11, color: { argb: 'FFFFFF' } };
        titleCell3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell3.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(4).height = 20;

        const headers = [
            'ت', 'المهام والواجبات', 'عدد الموظفين القائمين بالخدمة', 'عدد الأعمال ضمن الخطة',
            'نسبة الإنجاز أو عدد المنجز', 'القسم', 'الشعبة', 'الجهة المستفيدة', 'الجهة الساندة', 'تاريخ التنفيذ'
        ];

        const headerRow = worksheet.getRow(5);
        headerRow.height = 30;
        headers.forEach((hText, hIdx) => {
            const cell = headerRow.getCell(hIdx + 1);
            cell.value = hText;
            cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2A5298' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'CCCCCC' } },
                left: { style: 'thin', color: { argb: 'CCCCCC' } },
                bottom: { style: 'medium', color: { argb: '1E3C72' } },
                right: { style: 'thin', color: { argb: 'CCCCCC' } }
            };
        });

        submission.rows.forEach((row, rIdx) => {
            const dataRow = worksheet.getRow(6 + rIdx);
            dataRow.height = 25;
            const rowValues = [
                rIdx + 1,
                row.tasks,
                Number(row.employees) || row.employees || '',
                Number(row.planned) || row.planned || '',
                row.completion,
                row.section,
                row.department,
                row.beneficiary,
                row.supporting,
                row.date
            ];
            rowValues.forEach((val, valIdx) => {
                const cell = dataRow.getCell(valIdx + 1);
                cell.value = val;
                cell.font = { name: 'Arial', size: 10 };
                cell.alignment = { horizontal: valIdx === 1 ? 'right' : 'center', vertical: 'middle', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'E5E5E5' } },
                    left: { style: 'thin', color: { argb: 'E5E5E5' } },
                    bottom: { style: 'thin', color: { argb: 'E5E5E5' } },
                    right: { style: 'thin', color: { argb: 'E5E5E5' } }
                };
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent(`Report_${submission.governorate}_${submission.month}.xlsx`));
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) {
        console.error('Error generating Excel receipt:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Excel: ' + e.message);
    }
});

// Delete individual submission by employee (Public check name)
app.delete('/api/submissions/employee/:id', (req, res) => {
    try {
        const { name } = req.query;
        const subId = req.params.id;
        if (!name || name.trim() === '') {
            return res.status(400).json({ error: 'اسم المرسل مطلوب للتحقق من الصلاحية' });
        }

        const submission = db.getSubmissions().find(s => s.id === subId);
        if (!submission) {
            return res.status(404).json({ error: 'الاستمارة غير موجودة' });
        }

        // Verify ownership
        if (submission.userName.toLowerCase() !== name.trim().toLowerCase()) {
            return res.status(403).json({ error: 'لا تملك الصلاحية لحذف هذه الاستمارة' });
        }

        const success = db.deleteSubmission(subId);
        if (success) {
            res.json({ success: true });
        } else {
            res.status(500).json({ error: 'حدث خطأ أثناء الحذف' });
        }
    } catch (e) {
        res.status(500).json({ error: 'خطأ في حذف الاستمارة' });
    }
});

// Export cumulative Word for employee (Public)
app.get('/api/export/employee/cumulative/word', async (req, res) => {
    try {
        const { name, governorate, month } = req.query;
        if (!name || !governorate || !month) {
            return res.status(400).send('المدخلات غير كاملة');
        }

        // Fetch submissions for this employee name, governorate, and month
        const submissions = db.getSubmissions().filter(s => 
            s.userName.toLowerCase() === name.trim().toLowerCase() &&
            s.governorate === governorate &&
            s.month === month
        );

        if (submissions.length === 0) {
            return res.status(404).send('لا توجد بيانات لتصديرها');
        }

        // Merge rows
        let allRows = [];
        submissions.forEach(sub => {
            if (sub.rows && Array.isArray(sub.rows)) {
                allRows = allRows.concat(sub.rows);
            }
        });

        // Deduplicate rows
        const seen = new Set();
        const uniqueRows = allRows.filter(row => {
            const key = [
                (row.tasks || '').trim(),
                (row.employees || '').toString().trim(),
                (row.planned || '').toString().trim(),
                (row.completion || '').trim(),
                (row.section || '').trim(),
                (row.department || '').trim(),
                (row.beneficiary || '').trim(),
                (row.supporting || '').trim(),
                (row.date || '').trim()
            ].join('|||').toLowerCase();
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });

        const { Table, TableRow, TableCell, Document, Paragraph, TextRun, AlignmentType, WidthType, Packer } = docx;
        const docChildren = [];

        // Header block
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 50 }, children: [new TextRun({ text: 'قسم التصاريح الأمنية', bold: true, color: '1E3C72', size: 24 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 50 }, children: [new TextRun({ text: 'شعبة المتابعة', bold: true, color: '1E3C72', size: 18 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 50 }, children: [new TextRun({ text: 'استمارة الأنشطة الشهرية التراكمية', bold: true, color: '1E3C72', size: 20 })] }));
        docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 }, children: [new TextRun({ text: `محافظة (${governorate}) - لشهر (${month})`, bold: true, color: '2A5298', size: 16 })] }));

        const tableRows = [];
        const headerRow = new TableRow({
            children: [
                new TableCell({ width: { size: 400, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'ت', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'المهام والواجبات', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الموظفين', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 900, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'عدد الأعمال', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'نسبة الإنجاز', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'القسم', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الشعبة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة المستفيدة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'الجهة الساندة', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: '1E3C72' }, children: [new Paragraph({ children: [new TextRun({ text: 'تاريخ التنفيذ', bold: true, color: 'FFFFFF', size: 16 })], alignment: AlignmentType.CENTER })] })
            ]
        });
        tableRows.push(headerRow);

        uniqueRows.forEach((r, rIdx) => {
            const cells = [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(rIdx + 1), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.tasks || '', size: 15 })], alignment: AlignmentType.RIGHT })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.employees || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(r.planned || ''), size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.completion || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.section || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.department || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.beneficiary || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.supporting || '', size: 15 })], alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.date || '', size: 15 })], alignment: AlignmentType.CENTER })] })
            ];
            tableRows.push(new TableRow({ children: cells }));
        });

        const wordTable = new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
        docChildren.push(wordTable);

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 800, bottom: 800, left: 800, right: 800 }, size: { orientation: 'landscape' } } },
                children: docChildren
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent(`Report_Cumulative_${governorate}_${month}.docx`));
        res.send(buffer);
    } catch (e) {
        console.error('Error generating cumulative Word for employee:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Word: ' + e.message);
    }
});

// Export cumulative Excel for employee (Public)
app.get('/api/export/employee/cumulative/excel', async (req, res) => {
    try {
        const { name, governorate, month } = req.query;
        if (!name || !governorate || !month) {
            return res.status(400).send('المدخلات غير كاملة');
        }

        const submissions = db.getSubmissions().filter(s => 
            s.userName.toLowerCase() === name.trim().toLowerCase() &&
            s.governorate === governorate &&
            s.month === month
        );

        if (submissions.length === 0) {
            return res.status(404).send('لا توجد بيانات لتصديرها');
        }

        // Merge rows
        let allRows = [];
        submissions.forEach(sub => {
            if (sub.rows && Array.isArray(sub.rows)) {
                allRows = allRows.concat(sub.rows);
            }
        });

        // Deduplicate rows
        const seen = new Set();
        const uniqueRows = allRows.filter(row => {
            const key = [
                (row.tasks || '').trim(),
                (row.employees || '').toString().trim(),
                (row.planned || '').toString().trim(),
                (row.completion || '').trim(),
                (row.section || '').trim(),
                (row.department || '').trim(),
                (row.beneficiary || '').trim(),
                (row.supporting || '').trim(),
                (row.date || '').trim()
            ].join('|||').toLowerCase();
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet(governorate, { views: [{ rtl: true }] });

        worksheet.columns = [
            { key: 'seq', width: 6 },
            { key: 'tasks', width: 45 },
            { key: 'employees_count', width: 22 },
            { key: 'planned', width: 22 },
            { key: 'completion', width: 22 },
            { key: 'section', width: 18 },
            { key: 'dept', width: 18 },
            { key: 'beneficiary', width: 20 },
            { key: 'supporting', width: 20 },
            { key: 'exec_date', width: 15 }
        ];

        worksheet.mergeCells(1, 1, 1, 10);
        const titleCell1 = worksheet.getCell(1, 1);
        titleCell1.value = 'قسم التصاريح الأمنية';
        titleCell1.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFF' } };
        titleCell1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell1.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 25;

        worksheet.mergeCells(2, 1, 2, 10);
        const titleCellSub1 = worksheet.getCell(2, 1);
        titleCellSub1.value = 'شعبة المتابعة';
        titleCellSub1.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        titleCellSub1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCellSub1.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(2).height = 20;

        worksheet.mergeCells(3, 1, 3, 10);
        const titleCell2 = worksheet.getCell(3, 1);
        titleCell2.value = 'استمارة الأنشطة الشهرية التراكمية';
        titleCell2.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        titleCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell2.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(3).height = 20;

        worksheet.mergeCells(4, 1, 4, 10);
        const titleCell3 = worksheet.getCell(4, 1);
        titleCell3.value = `محافظة (${governorate}) - لشهر (${month})`;
        titleCell3.font = { name: 'Arial', size: 11, color: { argb: 'FFFFFF' } };
        titleCell3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3C72' } };
        titleCell3.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(4).height = 20;

        const headers = [
            'ت', 'المهام والواجبات', 'عدد الموظفين القائمين بالخدمة', 'عدد الأعمال ضمن الخطة',
            'نسبة الإنجاز أو عدد المنجز', 'القسم', 'الشعبة', 'الجهة المستفيدة', 'الجهة الساندة', 'تاريخ التنفيذ'
        ];

        const headerRow = worksheet.getRow(5);
        headerRow.height = 30;
        headers.forEach((hText, hIdx) => {
            const cell = headerRow.getCell(hIdx + 1);
            cell.value = hText;
            cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2A5298' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'CCCCCC' } },
                left: { style: 'thin', color: { argb: 'CCCCCC' } },
                bottom: { style: 'medium', color: { argb: '1E3C72' } },
                right: { style: 'thin', color: { argb: 'CCCCCC' } }
            };
        });

        uniqueRows.forEach((row, rIdx) => {
            const dataRow = worksheet.getRow(6 + rIdx);
            dataRow.height = 25;
            const rowValues = [
                rIdx + 1,
                row.tasks,
                Number(row.employees) || row.employees || '',
                Number(row.planned) || row.planned || '',
                row.completion,
                row.section,
                row.department,
                row.beneficiary,
                row.supporting,
                row.date
            ];
            rowValues.forEach((val, valIdx) => {
                const cell = dataRow.getCell(valIdx + 1);
                cell.value = val;
                cell.font = { name: 'Arial', size: 10 };
                cell.alignment = { horizontal: valIdx === 1 ? 'right' : 'center', vertical: 'middle', wrapText: true };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'E5E5E5' } },
                    left: { style: 'thin', color: { argb: 'E5E5E5' } },
                    bottom: { style: 'thin', color: { argb: 'E5E5E5' } },
                    right: { style: 'thin', color: { argb: 'E5E5E5' } }
                };
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent(`Report_Cumulative_${governorate}_${month}.xlsx`));
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) {
        console.error('Error generating cumulative Excel for employee:', e);
        res.status(500).send('خطأ أثناء إنشاء ملف Excel: ' + e.message);
    }
});

// Start Server
if (!process.env.VERCEL) {
    async function startServer() {
        try {
            await db.initDb();
            app.listen(PORT, () => {
                console.log(`Server is running on port ${PORT}`);
            });
        } catch (err) {
            console.error('Failed to initialize database on startup:', err);
            process.exit(1);
        }
    }
    startServer();
}

module.exports = app;
