const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const dotenv = require('dotenv');
const cors = require('cors');
const app = express();

dotenv.config();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(cors({
  origin: "*", // Change this to your frontend URL
  credentials: true,
}));

// Utility: Write title in A1 and start data from A2
const writeDataWithTitle = (filePath, data, title) => {
  const ws = xlsx.utils.json_to_sheet(data, { origin: "A2" }); // Start headers from A2
  ws['A1'] = { v: title }; // Write title at A1
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
  xlsx.writeFile(wb, filePath);
};

// Utility: Read data skipping the title row
const readDataSkippingTitle = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return {
    workbook,
    sheet,
    data: xlsx.utils.sheet_to_json(sheet, { range: 1, defval: '' }) // Skip A1
  };
};

// Create Excel files with title if they don't exist
const checkAndCreateFile = () => {
  if (!fs.existsSync('users.xlsx')) {
    const header = [{ username: '', email: '', password: '' }];
    writeDataWithTitle('users.xlsx', header, 'Users Data');
    console.log("Users Excel file created!");
  }

  if (!fs.existsSync('subscriptions.xlsx')) {
    const header = [{ email: '' }];
    const ws = xlsx.utils.json_to_sheet(header);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    xlsx.writeFile(wb, 'subscriptions.xlsx');
    console.log("Subscriptions Excel file created!");
  }

  if (!fs.existsSync('rentals.xlsx')) {
    const header = [{ name: '', email: '', phone: '', category: '', duration: '' }];
    writeDataWithTitle('rentals.xlsx', header, 'Rent User Data');
    console.log("Rentals Excel file created!");
  }

  if (!fs.existsSync('purchases.xlsx')) {
    const header = [{ name: '', email: '', phone: '', category: '', quantity: '' }];
    writeDataWithTitle('purchases.xlsx', header, 'Buy Users Data');
    console.log("Purchases Excel file created!");
  }
};

checkAndCreateFile();

// SIGNUP Route
app.post('/api/signup', async (req, res) => {
  const { username, email, password } = req.body;
  if (!username || !email || !password) {
    return res.status(400).json({ message: 'All fields are required' });
  }

  const { data } = readDataSkippingTitle('users.xlsx');

  if (data.find(u => u.email === email)) {
    return res.status(400).json({ message: 'Email already exists!' });
  }
  if (data.find(u => u.username === username)) {
    return res.status(400).json({ message: 'Username already exists!' });
  }

  const newUser = { username, email, password };
  data.push(newUser);

  writeDataWithTitle('users.xlsx', data, 'Users Data');
  res.status(200).json({ user: newUser, message: 'User signed up successfully!' });
});

// LOGIN Route
app.post('/api/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).json({ message: 'All fields are required' });
  }

  const { data } = readDataSkippingTitle('users.xlsx');
  const user = data.find(u => u.email === email && u.password === password);

  if (!user) {
    return res.status(400).json({ message: 'Invalid credentials!' });
  }

  // Just return the user's email as part of the response
  res.status(200).json({ message: 'Login successful!', email: user.email });
});

// SUBSCRIBE Route
app.post('/api/subscribe', (req, res) => {
  const { email } = req.body;
  if (!email) {
    return res.status(400).json({ message: 'Email is required to subscribe!' });
  }

  const workbook = xlsx.readFile('subscriptions.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);

  if (data.find(sub => sub.email === email)) {
    return res.status(400).json({ message: 'This email is already subscribed!' });
  }

  data.push({ email });

  const newSheet = xlsx.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  xlsx.writeFile(workbook, 'subscriptions.xlsx');

  res.status(200).json({ message: 'Subscription successful!' });
});

// RENT Route
app.post('/api/rent', (req, res) => {
  const { name, email, phone, category, duration } = req.body;

  if (!name || !email || !phone || !category || !duration) {
    return res.status(400).json({ message: 'All fields are required!' });
  }

  const { data } = readDataSkippingTitle('rentals.xlsx');

  const newRental = { name, email, phone, category, duration };
  data.push(newRental);

  writeDataWithTitle('rentals.xlsx', data, 'Rent User Data');

  res.status(200).json({ message: 'Rental request submitted successfully!', rental: newRental });
});

// BUY Route
app.post('/api/buy', (req, res) => {
  const { name, email, phone, category, quantity } = req.body;

  if (!name || !email || !phone || !category || !quantity) {
    return res.status(400).json({ message: 'All fields are required!' });
  }

  const { data } = readDataSkippingTitle('purchases.xlsx');

  const newPurchase = { name, email, phone, category, quantity };
  data.push(newPurchase);

  writeDataWithTitle('purchases.xlsx', data, 'Buy Users Data');

  res.status(200).json({ message: 'Purchase submitted successfully!', purchase: newPurchase });
});

// PROFILE Route (GET) - Fetch user profile using email in query params
app.get('/api/profile', (req, res) => {
  const { email } = req.query; // Expecting email as a query parameter

  if (!email) {
    return res.status(400).json({ message: 'Email is required!' });
  }

  const { data } = readDataSkippingTitle('users.xlsx');
  const user = data.find(u => u.email === email);

  if (!user) {
    return res.status(404).json({ message: 'User not found!' });
  }

  res.status(200).json({
    name: user.username,
    email: user.email
  });
});

// ORDERS Route (GET) - Fetch user orders (rentals and purchases) using email in query params
app.get('/api/orders', (req, res) => {
  const { email } = req.query; // Expecting email as a query parameter

  if (!email) {
    return res.status(400).json({ message: 'Email is required!' });
  }

  const { data: rentals } = readDataSkippingTitle('rentals.xlsx');
  const { data: purchases } = readDataSkippingTitle('purchases.xlsx');
  const userRentals = rentals.filter(r => r.email === email);
  const userPurchases = purchases.filter(p => p.email === email);
  res.status(200).json({
    rentals: userRentals,
    purchases: userPurchases
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
