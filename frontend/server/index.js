import express from 'express';
import cors from 'cors';
import sqlite3 from 'sqlite3';
import { open } from 'sqlite';

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Database setup
let db;
async function initializeDB() {
  db = await open({
    filename: './database.sqlite',
    driver: sqlite3.Database
  });

  await db.exec(`
    CREATE TABLE IF NOT EXISTS teams (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      team_name TEXT,
      team_size INTEGER
    );
    CREATE TABLE IF NOT EXISTS members (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      team_id INTEGER,
      name TEXT,
      college TEXT,
      year TEXT,
      branch TEXT,
      FOREIGN KEY(team_id) REFERENCES teams(id)
    );
    CREATE TABLE IF NOT EXISTS masterclass_registrations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT,
      college TEXT,
      year TEXT,
      branch TEXT
    );
  `);
  console.log('Database initialized');
}

initializeDB();

// Routes
app.get('/', (req, res) => {
  res.send('AIFi Backend Running');
});

// Hackathon Registration
app.post('/api/register-hackathon', async (req, res) => {
  const { teamName, teamSize, members } = req.body;
  try {
    const result = await db.run(
      'INSERT INTO teams (team_name, team_size) VALUES (?, ?)',
      [teamName, teamSize]
    );
    const teamId = result.lastID;

    for (const member of members) {
      await db.run(
        'INSERT INTO members (team_id, name, college, year, branch) VALUES (?, ?, ?, ?, ?)',
        [teamId, member.name, member.college, member.year, member.branch]
      );
    }
    res.json({ success: true, message: 'Team registered successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: 'Registration failed' });
  }
});

// Masterclass Registration
app.post('/api/register-masterclass', async (req, res) => {
  const { name, college, year, branch } = req.body;
  try {
    await db.run(
      'INSERT INTO masterclass_registrations (name, college, year, branch) VALUES (?, ?, ?, ?)',
      [name, college, year, branch]
    );
    res.json({ success: true, message: 'Registered for Masterclass successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: 'Registration failed' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
