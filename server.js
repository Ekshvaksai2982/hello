const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

// Enable CORS
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Function to read Excel data
function readExcelData(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    return jsonData;
  } catch (error) {
    console.error('Error reading Excel file:', error.message);
    return [];
  }
}

// Function to append data to Excel
function appendToExcel(filePath, data) {
  try {
    console.log('Appending data to:', filePath, 'Data:', data);
    let workbook;
    let worksheet;

    if (fs.existsSync(filePath)) {
      console.log('Reading existing file:', filePath);
      workbook = xlsx.readFile(filePath);
      worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
      console.log('Creating new file:', filePath);
      workbook = xlsx.utils.book_new();
      worksheet = xlsx.utils.json_to_sheet([]);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'UserData');
      xlsx.utils.sheet_add_json(worksheet, [{
        'First Name': '',
        'Email': '',
        'Phone': '',
        'Signup Date': '',
        'Signup Time': '',
        'Job Role': '',
        'Probability Score': '',
        'Resume File': ''
      }], { skipHeader: false, origin: 'A1' });
    }

    const existingData = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
    existingData.push(data);
    const newWorksheet = xlsx.utils.json_to_sheet(existingData, { header: [
      'First Name', 'Email', 'Phone', 'Signup Date', 'Signup Time', 'Job Role', 'Probability Score', 'Resume File'
    ] });
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    console.log('Writing to file:', filePath);
    xlsx.writeFileSync(workbook, filePath);
    console.log('Data appended successfully');
  } catch (error) {
    console.error('Error appending to Excel:', error.message);
    throw error;
  }
}

// Simple route
app.get('/', (req, res) => {
  res.send('Hello, world!');
});

// Handle file uploads and analyze resume
app.post('/upload', upload.single('resume'), async (req, res) => {
  console.log('Request body:', req.body);
  console.log('File received:', req.file);
  const { firstName, email, phone, jobRole } = req.body;

  if (!req.file) {
    console.error('No file uploaded');
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const filePath = path.join(__dirname, req.file.path);
  try {
    const dataBuffer = fs.readFileSync(filePath);
    const resumeFileName = req.file.originalname;
    const data = await pdfParse(dataBuffer);
    const resumeText = data.text;

    const excelData = readExcelData(path.join(__dirname, 'jobrolespskillsframeworks.xlsx'));
    const analysisResult = analyzeResume(resumeText, jobRole, excelData);

    const response = {
      jobRole: jobRole,
      probability: analysisResult.probability,
      additionalSkills: analysisResult.additionalSkills,
      additionalFrameworks: analysisResult.additionalFrameworks,
      feedback: analysisResult.feedback,
    };

    const now = new Date();
    const signupDate = now.toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
    const signupTime = now.toLocaleTimeString('en-US', { hour12: false });

    const excelDataRow = {
      'First Name': firstName || 'N/A',
      'Email': email || 'N/A',
      'Phone': phone || 'N/A',
      'Signup Date': signupDate,
      'Signup Time': signupTime,
      'Job Role': jobRole || 'N/A',
      'Probability Score': analysisResult.probability,
      'Resume File': resumeFileName
    };

    const excelFilePath = path.join(__dirname, 'userdata.xlsx');
    appendToExcel(excelFilePath, excelDataRow);

    fs.unlinkSync(filePath);
    res.json(response);
  } catch (error) {
    console.error('Error processing request:', error.message);
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    res.status(500).json({ error: 'Error processing request: ' + error.message });
  }
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Server error:', err.stack);
  res.status(500).send('Something went wrong!');
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});

// Analyze resume text (unchanged)
function analyzeResume(resumeText, jobRole, excelData) {
  const jobData = excelData.find(item => item['JOB ROLES'] === jobRole);
  if (!jobData) {
    return {
      jobRole,
      probability: 0,
      additionalSkills: 'Job role not found in the dataset',
      additionalFrameworks: 'Job role not found in the dataset',
      feedback: 'Job role not found in the dataset',
    };
  }

  const requiredSkills = jobData['PROGRAMMING SKILLS'].split(',').map(skill => skill.trim());
  const requiredFrameworks = jobData['FRAMEWORKS'].split(',').map(framework => framework.trim());
  const skillsFound = [];
  const frameworksFound = [];
  const additionalSkills = [];
  const additionalFrameworks = [];
  let probability = 0;
  let feedback = 'Better luck next time. Consider improving your skills in certain areas.';

  requiredSkills.forEach(skill => {
    if (resumeText.toLowerCase().includes(skill.toLowerCase())) {
      skillsFound.push(skill);
    } else {
      additionalSkills.push(skill);
    }
  });

  requiredFrameworks.forEach(framework => {
    if (resumeText.toLowerCase().includes(framework.toLowerCase())) {
      frameworksFound.push(framework);
    } else {
      additionalFrameworks.push(framework);
    }
  });

  const skillsProbability = (skillsFound.length / requiredSkills.length) * 50;
  const frameworksProbability = (frameworksFound.length / requiredFrameworks.length) * 50;
  probability = skillsProbability + frameworksProbability;

  if (probability === 100) {
    feedback = 'Great job! You are a perfect match for this role!';
  } else if (probability >= 50) {
    feedback = 'You have some of the required skills and frameworks. Consider improving: ' + additionalSkills.join(', ') + ', ' + additionalFrameworks.join(', ');
  } else {
    feedback = 'You need to improve your skills and frameworks significantly. Consider learning: ' + additionalSkills.join(', ') + ', ' + additionalFrameworks.join(', ');
  }

  return {
    jobRole,
    probability,
    additionalSkills: additionalSkills.join(', ') || 'None',
    additionalFrameworks: additionalFrameworks.join(', ') || 'None',
    feedback,
  };
}