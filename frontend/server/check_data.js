import sqlite3 from 'sqlite3';
import { open } from 'sqlite';

async function checkData() {
    const db = await open({
        filename: './database.sqlite',
        driver: sqlite3.Database
    });

    console.log('--- Teams ---');
    const teams = await db.all('SELECT * FROM teams');
    console.table(teams);

    console.log('\n--- Members ---');
    const members = await db.all('SELECT * FROM members');
    console.table(members);

    console.log('\n--- Masterclass Registrations ---');
    const registrations = await db.all('SELECT * FROM masterclass_registrations');
    console.table(registrations);
}

checkData().catch(console.error);
