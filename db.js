const fs = require('fs');
const path = require('path');

const DB_DIR = process.env.DB_DIR || __dirname;
const DB_PATH = path.join(DB_DIR, 'db.json');

// In-memory cache
let dbCache = null;
let isSaving = false;
let needsSave = false;
let initPromise = null;

const FIREBASE_URL = process.env.FIREBASE_URL || 'https://inshetaa-default-rtdb.firebaseio.com/db.json';

// Helper to read database
function readDb() {
    if (!dbCache) {
        if (fs.existsSync(DB_PATH)) {
            try {
                const data = fs.readFileSync(DB_PATH, 'utf8');
                dbCache = JSON.parse(data);
            } catch (e) {
                dbCache = { users: [], submissions: [] };
            }
        } else {
            dbCache = { users: [], submissions: [] };
        }
    }
    return dbCache;
}

// Background sync to Firebase
function triggerCloudSave() {
    if (!FIREBASE_URL) return;
    
    if (isSaving) {
        needsSave = true;
        return;
    }

    isSaving = true;
    needsSave = false;

    fetch(FIREBASE_URL, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(dbCache)
    })
    .then(res => {
        if (!res.ok) {
            console.error('Failed to save to Firebase, status:', res.status);
        }
    })
    .catch(err => {
        console.error('Error saving to Firebase:', err);
    })
    .finally(() => {
        isSaving = false;
        if (needsSave) {
            triggerCloudSave();
        }
    });
}

// Helper to write database
function writeDb(data) {
    dbCache = data;
    try {
        fs.writeFile(DB_PATH, JSON.stringify(dbCache, null, 4), 'utf8', () => {});
    } catch (error) {
        console.error('Error writing local database backup:', error);
    }
    triggerCloudSave();
    return true;
}

// Initialize database (Async for cloud support)
async function initDb() {
    if (dbCache) return;
    if (initPromise) return initPromise;

    initPromise = (async () => {
        if (FIREBASE_URL) {
            console.log('Connecting to Firebase Realtime Database...');
            try {
                const response = await fetch(FIREBASE_URL);
                if (response.ok) {
                    const data = await response.json();
                    if (data && data.users && data.submissions) {
                        dbCache = data;
                        console.log('✅ Loaded database from Firebase successfully.');
                        return;
                    }
                }
                console.log('⚠️ Firebase database empty or invalid, initializing defaults.');
            } catch (error) {
                console.error('Error connecting to Firebase, falling back to local file:', error);
            }
        }

        if (fs.existsSync(DB_PATH)) {
            try {
                const data = fs.readFileSync(DB_PATH, 'utf8');
                dbCache = JSON.parse(data);
                console.log('✅ Loaded database from local db.json.');
                return;
            } catch (e) {
                console.error('Error loading local db.json:', e);
            }
        }

        dbCache = {
            users: [
                {
                    id: 'admin_default',
                    name: 'مدير النظام',
                    passcode: 'admin123',
                    department: 'إدارة النظام',
                    section: 'القسم الأمني',
                    role: 'admin'
                }
            ],
            submissions: []
        };
        writeDb(dbCache);
        console.log('✅ Initialized default database.');
    })();

    return initPromise;
}

// User operations
function getUsers() {
    const db = readDb();
    return db.users;
}

function addUser(name, passcode, department, section) {
    const db = readDb();
    // Check if passcode is already used
    if (db.users.some(u => u.passcode === passcode)) {
        throw new Error('رمز الدخول مستخدم بالفعل لموظف آخر');
    }
    const newUser = {
        id: 'user_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
        name,
        passcode,
        department,
        section,
        role: 'employee'
    };
    db.users.push(newUser);
    writeDb(db);
    return newUser;
}

function deleteUser(userId) {
    const db = readDb();
    const index = db.users.findIndex(u => u.id === userId);
    if (index === -1) return false;
    if (db.users[index].role === 'admin') {
        throw new Error('لا يمكن حذف حساب مدير النظام الرئيسي');
    }
    db.users.splice(index, 1);
    writeDb(db);
    return true;
}

function findUserByPasscode(passcode) {
    const db = readDb();
    return db.users.find(u => u.passcode === passcode);
}

// Submission operations
function getSubmissions() {
    const db = readDb();
    return db.submissions;
}

function addSubmission(userId, userName, governorate, month, rows) {
    const db = readDb();
    const newSubmission = {
        id: 'sub_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
        userId,
        userName,
        governorate,
        month,
        rows,
        submissionDate: new Date().toISOString()
    };
    db.submissions.push(newSubmission);
    writeDb(db);
    return newSubmission;
}

function deleteSubmission(subId) {
    const db = readDb();
    const index = db.submissions.findIndex(s => s.id === subId);
    if (index === -1) return false;
    db.submissions.splice(index, 1);
    writeDb(db);
    return true;
}

function updateAdminPasscode(newPasscode) {
    const db = readDb();
    const admin = db.users.find(u => u.role === 'admin');
    if (!admin) {
        throw new Error('لم يتم العثور على حساب مدير النظام');
    }
    admin.passcode = newPasscode;
    writeDb(db);
    return true;
}

module.exports = {
    initDb,
    getUsers,
    addUser,
    deleteUser,
    findUserByPasscode,
    getSubmissions,
    addSubmission,
    deleteSubmission,
    updateAdminPasscode
};
