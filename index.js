import { Command } from 'commander';
import inquirer from 'inquirer';
import fs from 'fs';
import chalk from 'chalk';
import xlsx from 'xlsx';

import axios from 'axios';



const program = new Command();
const filePath = './coursess.json';
const studentsFilePath = './students.json';
const excelFilePath = './output.xlsx'; // Path to save the Excel file
const excelFilePathStudent = './Studentoutput.xlsx'; // Path to save the Excel file

const BOT_TOKEN = ''; //Put your Bot Token
const CHAT_ID = ''; //Put Chat ID


//Function to send Message Via Telegram
function sendTelegramMessage(chatId, text) {
    const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;
    return axios.post(url, {
      chat_id: chatId,
      text: text,
    });
  }

  const v=new Date().toISOString().split('T')[0]
const q=new Date().toISOString().split('0')[1]
console.log(`v : ${v}`)
console.log(`q : ${q}`)
console.log("pass is q+v ")

// Set your desired password here
const correctPassword = q+new Date().toISOString().split('T')[0];

console.log(chalk.bold('Available Commands:'));
  program.commands.forEach((cmd) => {
    const alias = cmd.alias() ? `(${chalk.bgBlue(cmd.alias())})` : '';
    console.log(
      `  ${chalk.green(cmd.name())} ${alias} - ${chalk.yellow(cmd.description())}`
    );
  });


  // Function to prompt for password
async function askForPassword() {
    const answers = await inquirer.prompt([
      {
        type: 'password',
        name: 'password',
        message: 'Please enter the password:',
        mask: '*',
      },
    ]);
  
    if (answers.password !== correctPassword) {
      console.log(chalk.red('Incorrect password. Exiting program.'));
      process.exit(1); // Exit the program if password is incorrect
    } else {
      console.log(chalk.green('Password correct. Access granted.'));
    }
  }


//Questions for Courses

const questions = [
    {
      type: 'input',
      name: 'title',
      message: 'Please enter course title',
    },
    {
      type: 'number',
      name: 'price',
      message: 'Please enter course price',
    },
    {
      type: 'input',
      name: 'Course_ID',
      message: 'Please enter course ID',
    },
  ];




  //Question For Student

  const questions2 = [
    {
      type: 'input',
      name: 'Name',
      message: 'Enter Student name',
    },
    {
      type: 'input',
      name: 'Phone_Number',
      message: 'Please enter Phone Number',
    },
    {
      type: 'input',
      name: 'Student_ID',
      message: 'Please enter Student ID',
    },

  ];

  program
  .name('Code-Courses-Manager')
  .description('CLI to manage courses')
  .version('12.1.0');

 //Registration
 program
 .command('registration')
 .alias('r')
 .description('Add new student and register courses')
 .action(() => {
   if (!fs.existsSync(filePath)) {
     console.log('Courses file not found. Please add courses first.');
     return;
   }

   const courses = JSON.parse(fs.readFileSync(filePath, 'utf8'));

   if (courses.length === 0) {
     console.log('No courses available. Please add courses first.');
     return;
   }

   const courseChoices = courses.map((course) => ({
     name: `${course.title} (ID: ${course.Course_ID}, Price: ${course.price})`,
     value: course.Course_ID,
   }));

   inquirer.prompt(questions2).then((studentAnswers) => {
     // Load existing students
     const existingStudents = fs.existsSync(studentsFilePath)
       ? JSON.parse(fs.readFileSync(studentsFilePath, 'utf8'))
       : [];

     // Check for unique Student_ID
     const isDuplicateID = existingStudents.some(
       (student) => student.Student_ID === studentAnswers.Student_ID
     );

     if (isDuplicateID) {
       console.log(`Error: Student ID ${studentAnswers.Student_ID} already exists. Please try again with a unique ID.`);
       return;
     }

     inquirer
       .prompt([
         {
           type: 'checkbox',
           name: 'selectedCourses',
           message: 'Select courses to register:',
           choices: courseChoices,
         },
       ])
       .then(({ selectedCourses }) => {
         if (selectedCourses.length === 0) {
           console.log('No courses selected.');
           return;
         }

         const selectedCourseDetails = courses.filter((course) =>
           selectedCourses.includes(course.Course_ID)
         );

         const totalPrice = selectedCourseDetails.reduce(
           (total, course) => total + course.price,
           0
         );
         console.log('Total Price:', totalPrice);

         inquirer
           .prompt([
             {
               type: 'number',
               name: 'firstPayment',
               message: `The total price is ${totalPrice}. Enter the first payment amount:`,
               validate: (value) => {
                 if (value <= 0 || value > totalPrice) {
                   return `Please enter a value between 1 and ${totalPrice}.`;
                 }
                 return true;
               },
             },
           ])
           .then(({ firstPayment }) => {
             const remainingBalance = totalPrice - firstPayment;
             const paymentDate = new Date().toISOString().split('T')[0]; // Extract date part only

             // Convert the selected courses into a readable string
             const registeredCoursesString = selectedCourseDetails
               .map((course) => course.title)
               .join(', ');

             studentAnswers.registeredCourses = registeredCoursesString; // Store as a string
             studentAnswers.totalPrice = totalPrice;
             studentAnswers.firstPayment = firstPayment;
             studentAnswers.remainingBalance = remainingBalance;
             studentAnswers.paymentDate = paymentDate;

             existingStudents.push(studentAnswers);
             fs.writeFileSync(
               studentsFilePath,
               JSON.stringify(existingStudents, null, 2),
               'utf8'
             );

             console.log('Student registered successfully!');
             console.log('Selected Courses:', registeredCoursesString);
             console.log('First Payment:', firstPayment);
             console.log('Remaining Balance:', remainingBalance);
             console.log('Payment Date:', paymentDate);



             const message = `
📞 **From: Zero To Hero
👤 **To: ${studentAnswers.Name}
📱 (${studentAnswers.Phone_Number})

Ure ID is ${studentAnswers.Student_ID}
🎓 **Course Registration Confirmation** 🎉

📚 **Courses Registered:
- ${registeredCoursesString}

💰 **Payment Details:
- 💵 **Total Price: ${totalPrice}
- 💳 **First Payment: ${firstPayment}
- 🏷️ **Remaining Balance: ${remainingBalance}

🗓️ **Date: ${studentAnswers.paymentDate}

🎉 Thank you for registering and best of luck with your courses!
📘 Stay focused and succeed! 🚀
`;

             sendTelegramMessage(CHAT_ID, message)
               .then(() => console.log('Message sent successfully!'))
               .catch((err) => console.error('Error sending message:', err));
           });
       });
   });
 });



// Second Payment
program
.command('second-payment')
.alias('sp')
.description('Process second payment for a student')
.action(() => {
  if (!fs.existsSync(studentsFilePath)) {
    console.log('No students found. Please register students first.');
    return;
  }

  const students = JSON.parse(fs.readFileSync(studentsFilePath, 'utf8'));

  inquirer
    .prompt([
      {
        type: 'input',
        name: 'Student_ID',
        message: 'Enter the Student ID:',
        validate: (value) => {
          if (!students.find((student) => student.Student_ID === value)) {
            return 'Student not found. Please enter a valid Student ID.';
          }
          return true;
        },
      },
    ])
    .then(({ Student_ID }) => {
      const student = students.find((s) => s.Student_ID === Student_ID);

      console.log(`Student Found: ${student.Name}`);
      console.log(`Current Remaining Balance: ${student.remainingBalance}`);

      inquirer
        .prompt([
          {
            type: 'number',
            name: 'secondPayment',
            message: `Enter the second payment amount (max: ${student.remainingBalance}):`,
            validate: (value) => {
              if (value < 0 || value > student.remainingBalance) {
                return `Please enter a value between 0 and ${student.remainingBalance}.`;
              }
              return true;
            },
          },
        ])
        .then(({ secondPayment }) => {
          student.secondPayment = secondPayment;
          student.remainingBalance -= secondPayment;
          const secondPaymentDate = new Date().toISOString().split('T')[0];
student.secondPaymentDate = secondPaymentDate;

          // Update students file
          fs.writeFileSync(
            studentsFilePath,
            JSON.stringify(students, null, 2),
            'utf8'
          );

          console.log('Second payment processed successfully!');
          console.log(`Updated Remaining Balance: ${student.remainingBalance}`);


          const message = `
          📞 **From: Zero To Hero
          👤 **To: ${student.Name}

          📱 (${student.Phone_Number})

       💰 Updated Remaining Balance: ${student.remainingBalance}

         🎉 Thank you for your trust
        📘 Stay focused and succeed! 🚀
        🗓️ date: ${v}


          `;

                        sendTelegramMessage(CHAT_ID, message)
                          .then(() => console.log('Message sent successfully!'))
                          .catch((err) => console.error('Error sending message:', err));
        });
    });
});


 //To Disable Student List

 program
 .command('Student')
 .alias('S')
 .description('List all Student')
 .action(() => {
   if (fs.existsSync(studentsFilePath)) {
     const fileContent = fs.readFileSync(studentsFilePath, 'utf8');
     const Student = JSON.parse(fileContent);
     console.log('Student:', Student);
   } else {
     console.log('No Student found!');
   }
 });







 // Find and print student details by name

 program
 .command('findStudent')
 .alias('f')
 .description('Find and display details of a student by name')
 .action(() => {
   inquirer
     .prompt([
       {
         type: 'input',
         name: 'Name',
         message: 'Enter the name of the student to search:',
       },
     ])
     .then(({ Name }) => {
       if (fs.existsSync(studentsFilePath)) {
         const fileContent = fs.readFileSync(studentsFilePath, 'utf8');
         const students = fileContent ? JSON.parse(fileContent) : [];

        const student = students.find((s) => s.Name.toLowerCase() === Name.toLowerCase());

         if (student) {
           console.log('Student Details:', student);
         } else {
           console.log(`No student found with the name "${Name}".`);
         }
       } else {
         console.log('No students found!');
       }
     });
 });



         // Find and print student details by ID

         program
         .command('findStudentByID')
         .alias('fID')
         .description('Find and display details of a student by Student ID')
         .action(() => {
           inquirer
             .prompt([
               {
                 type: 'number',
                 name: 'Student_ID',
                 message: 'Enter the Student ID of the student to search:',
               },
             ])
             .then(({ Student_ID }) => {
               if (fs.existsSync(studentsFilePath)) {
                 const fileContent = fs.readFileSync(studentsFilePath, 'utf8');
                 const students = fileContent ? JSON.parse(fileContent) : [];
     
                 const student = students.find((s) => s.Student_ID === Student_ID);
     
                 if (student) {
                   console.log('Student Details:', student);
                 } else {
                   console.log(`No student found with the Student ID "${Student_ID}".`);
                 }
               } else {
                 console.log('No students found!');
               }
             });
         });
     
     
     
     
     
     
     
     
     
         // Add Courses to Student Profile
     
         program
       .command('addCourseToStudent')
       .alias('acs')
       .description('Add courses to an existing student profile')
       .action(() => {
         inquirer
           .prompt([
             {
               type: 'number',
               name: 'Student_ID',
               message: 'Enter the Student ID to add courses:',
             },
           ])
           .then(({ Student_ID }) => {
             if (fs.existsSync(studentsFilePath) && fs.existsSync(filePath)) {
               const students = JSON.parse(fs.readFileSync(studentsFilePath, 'utf8'));
               const courses = JSON.parse(fs.readFileSync(filePath, 'utf8'));
     
               const student = students.find((s) => s.Student_ID === Student_ID);
     
              if (!student) {
                 console.log(`No student found with ID: ${Student_ID}`);
                 return;
               }
     
               const courseChoices = courses.map((course) => ({
                 name: `${course.title} (ID: ${course.Course_ID}, Price: ${course.price})`,
                 value: course.Course_ID,
               }));
     
               inquirer
                 .prompt([
                   {
                     type: 'checkbox',
                     name: 'selectedCourses',
                     message: 'Select courses to add:',
                     choices: courseChoices,
                   },
                 ])
                 .then(({ selectedCourses }) => {
                   const selectedCourseDetails = courses.filter((course) =>
                     selectedCourses.includes(course.Course_ID)
                   );
     
                   student.registeredCourses.push(...selectedCourseDetails);
                   student.totalPrice += selectedCourseDetails.reduce(
                     (total, course) => total + course.price,
                     0
                   );
     
                   fs.writeFileSync(studentsFilePath, JSON.stringify(students, null, 2), 'utf8');
                   console.log('Courses added successfully!');
                   student.registeredCourses.push(...selectedCourseDetails);
       student.totalPrice += selectedCourseDetails.reduce(
         (total, course) => total + course.price,
         0
       );
     
                 });
             } else {
               console.log('Students or courses file not found.');
             }
           });
       });
     
     
     
     
     
     
     
     
     
       // Remove Courses from Student Profile
       program
       .command('removeCourseFromStudent')
       .alias('rcs')
       .description('Remove courses from an existing student profile')
       .action(() => {
         inquirer
           .prompt([
             {
               type: 'number',
               name: 'Student_ID',
               message: 'Enter the Student ID to remove courses:',
             },
           ])
           .then(({ Student_ID }) => {
             if (fs.existsSync(studentsFilePath)) {
               const students = JSON.parse(fs.readFileSync(studentsFilePath, 'utf8'));
               const student = students.find((s) => s.Student_ID === Student_ID);
     
               if (!student) {
                 console.log(`No student found with ID: ${Student_ID}`);
                 return;
               }
     
               const courseChoices = student.registeredCourses.map((course) => ({
                 name: `${course.title} (ID: ${course.Course_ID}, Price: ${course.price})`,
                 value: course.Course_ID,
               }));
     
               inquirer
                 .prompt([
                   {
                     type: 'checkbox',
                     name: 'selectedCourses',
                     message: 'Select courses to remove:',
                     choices: courseChoices,
                   },
                 ])
                 .then(({ selectedCourses }) => {
                   const updatedCourses = student.registeredCourses.filter(
                     (course) => !selectedCourses.includes(course.Course_ID)
                   );
     
                   const removedCourses = student.registeredCourses.filter((course) =>
                     selectedCourses.includes(course.Course_ID)
                   );
     
                   student.registeredCourses = updatedCourses;
                   student.totalPrice -= removedCourses.reduce(
                     (total, course) => total + course.price,
                     0
                   );

                   fs.writeFileSync(studentsFilePath, JSON.stringify(students, null, 2), 'utf8');
                   console.log('Courses removed successfully!');
                   student.registeredCourses = updatedCourses;
       student.totalPrice -= removedCourses.reduce(
         (total, course) => total + course.price,
         0
       );

                 });
             } else {
               console.log('Students file not found.');
             }
           });
       });

 //Remove Student with ID

 program
 .command('Remove')
 .alias('re')
 .description('Remove Student')
 .action(() => {
   inquirer
     .prompt([
       {
         type: 'input',
         name: 'Student_ID',
         message: 'Enter the Student ID to delete:',
       },
     ])
     .then(({ Student_ID }) => {
       if (fs.existsSync(studentsFilePath)) {
         const fileContent = fs.readFileSync(studentsFilePath, 'utf8');
         const Student = fileContent ? JSON.parse(fileContent) : [];

         const updatedStudent = Student.filter((Student) => Student.Student_ID !== Student_ID);

         if (Student.length === updatedStudent.length) {
           console.log(`No Student found with ID: ${Student_ID}`);
         } else {
           fs.writeFileSync(studentsFilePath, JSON.stringify(updatedStudent, null, 2), 'utf8');
           console.log(`Student with ID ${Student_ID} has been deleted.`);
         }
       } else {
         console.log('No Student found to delete.');
       }
     });
 });









 //To Add Courses


program
.command('add')
.alias('a')
.description('Add a course')
.action(() => {
 inquirer.prompt(questions).then((answers) => {
   if (fs.existsSync(filePath)) {
     const fileContent = fs.readFileSync(filePath, 'utf8');
     const fileContentJson = fileContent ? JSON.parse(fileContent) : [];
     fileContentJson.push(answers);
     fs.writeFileSync(filePath, JSON.stringify(fileContentJson, null, 2), 'utf8');
     console.log('Course added successfully!');
   } else {
     fs.writeFileSync(filePath, JSON.stringify([answers], null, 2), 'utf8');
     console.log('Course added successfully!');
   }
 });
});








//To Disable Courses List

program
.command('list')
.alias('l')
.description('List all courses')
.action(() => {
 if (fs.existsSync(filePath)) {
   const fileContent = fs.readFileSync(filePath, 'utf8');
   const courses = JSON.parse(fileContent);
   console.log('Courses:', courses);
 } else {
   console.log('No courses found!');
 }
});









//To Calculate Price

program
.command('calculate')
.alias('c')
.description('Calculate the total price for selected courses')
.action(() => {
 if (fs.existsSync(filePath)) {
   const fileContent = fs.readFileSync(filePath, 'utf8');
   const courses = JSON.parse(fileContent);

   if (courses.length === 0) {
     console.log('No courses available.');
     return;
   }

   const courseChoices = courses.map((course) => ({
     name: `${course.title} (ID: ${course.Course_ID}, Price: ${course.price})`,
     value: course.Course_ID,
   }));

   inquirer
     .prompt([
       {
         type: 'checkbox',
         name: 'selectedCourses',
         message: 'Select courses to register:',
         choices: courseChoices,
       },
     ])
     .then(({ selectedCourses }) => {
       if (selectedCourses.length === 0) {
         console.log('No courses selected.');
         return;
       }

       const selectedCourseDetails = courses.filter((course) =>
         selectedCourses.includes(course.Course_ID)
       );

       const totalPrice = selectedCourseDetails.reduce((total, course) => total + course.price, 0);

       console.log('Selected Courses:', selectedCourseDetails);
       console.log('Total Price:', totalPrice);

     });
 } else {
   console.log('No courses file found.');
 }
});










//To Delete Courses

program
.command('delete')
.alias('d')
.description('Delete a course by ID')
.action(() => {
 inquirer
   .prompt([
     {
       type: 'number',
       name: 'Course_ID',
       message: 'Enter the Course ID to delete:',
     },
   ])
   .then(({ Course_ID }) => {
     if (fs.existsSync(filePath)) {
       const fileContent = fs.readFileSync(filePath, 'utf8');
       const courses = fileContent ? JSON.parse(fileContent) : [];

       const updatedCourses = courses.filter((course) => course.Course_ID !== Course_ID);

       if (courses.length === updatedCourses.length) {
         console.log(`No course found with ID: ${Course_ID}`);
       } else {
         fs.writeFileSync(filePath, JSON.stringify(updatedCourses, null, 2), 'utf8');
         console.log(`Course with ID ${Course_ID} has been deleted.`);
       }
     } else {
       console.log('No courses found to delete.');
     }
   });
});










// Add Course Editing Command
program
.command('edit')
.alias('e')
.description('Edit an existing course')
.action(() => {
if (fs.existsSync(filePath)) {
const fileContent = fs.readFileSync(filePath, 'utf8');
const courses = JSON.parse(fileContent);

if (courses.length === 0) {
 console.log('No courses available to edit.');
 return;
}

// Prompt user to select a course to edit
const courseChoices = courses.map((course) => ({
 name: `${course.title} (ID: ${course.Course_ID})`,
 value: course.Course_ID,
}));

inquirer
 .prompt([
   {
     type: 'list',
     name: 'selectedCourse',
     message: 'Select a course to edit:',
     choices: courseChoices,
   },
 ])
 .then(({ selectedCourse }) => {
   // Find the selected course
   const courseToEdit = courses.find((course) => course.Course_ID === selectedCourse);

   if (!courseToEdit) {
     console.log('Course not found!');
     return;
   }

   // Prompt user to edit course details
   inquirer
     .prompt([
       {
         type: 'input',
         name: 'title',
         message: `Enter new title for ${courseToEdit.title} (current: ${courseToEdit.title}):`,
         default: courseToEdit.title,
       },
       {
         type: 'number',
         name: 'price',
         message: `Enter new price for ${courseToEdit.title} (current: ${courseToEdit.price}):`,
         default: courseToEdit.price,
       },
       {
         type: 'number',
         name: 'Course_ID',
         message: `Enter new Course ID for ${courseToEdit.title} (current: ${courseToEdit.Course_ID}):`,
         default: courseToEdit.Course_ID,
       },
     ])
     .then((updatedCourse) => {
       // Update the course with new details
       courseToEdit.title = updatedCourse.title;
       courseToEdit.price = updatedCourse.price;
       courseToEdit.Course_ID = updatedCourse.Course_ID;

       // Save updated courses to file
       fs.writeFileSync(filePath, JSON.stringify(courses, null, 2), 'utf8');
       console.log('Course updated successfully!');
     });
 });
} else {
console.log('No courses file found.');
}
});












// Add Student Editing Command
program
.command('editStudent')
.alias('es')
.description('Edit an existing student')
.action(() => {
if (fs.existsSync(studentsFilePath)) {
 const fileContent = fs.readFileSync(studentsFilePath, 'utf8');
 const students = JSON.parse(fileContent);

 if (students.length === 0) {
   console.log('No students available to edit.');
   return;
 }

 // Prompt user to select a student to edit
 const studentChoices = students.map((student) => ({
   name: `${student.Name} (ID: ${student.Student_ID})`,
   value: student.Student_ID,
 }));

 inquirer
   .prompt([
     {
       type: 'list',
       name: 'selectedStudent',
       message: 'Select a student to edit:',
       choices: studentChoices,
     },
   ])
   .then(({ selectedStudent }) => {
     // Find the selected student
     const studentToEdit = students.find((student) => student.Student_ID === selectedStudent);

     if (!studentToEdit) {
       console.log('Student not found!');
       return;
     }

     // Prompt user to edit student details
     inquirer
       .prompt([
         {
           type: 'input',
           name: 'Name',
           message: `Enter new name for ${studentToEdit.Name} (current: ${studentToEdit.Name}):`,
           default: studentToEdit.Name,
         },
         {
           type: 'number',
           name: 'Phone_Number',
           message: `Enter new phone number for ${studentToEdit.Name} (current: ${studentToEdit.Phone_Number}):`,
           default: studentToEdit.Phone_Number,
         },
         {
           type: 'number',
           name: 'Student_ID',
           message: `Enter new Student ID for ${studentToEdit.Name} (current: ${studentToEdit.Student_ID}):`,
           default: studentToEdit.Student_ID,
         },
       ])
       .then((updatedStudent) => {
         // Update the student with new details
         studentToEdit.Name = updatedStudent.Name;
         studentToEdit.Phone_Number = updatedStudent.Phone_Number;
         studentToEdit.Student_ID = updatedStudent.Student_ID;

         // Save updated students to file
         fs.writeFileSync(studentsFilePath, JSON.stringify(students, null, 2), 'utf8');
         console.log('Student updated successfully!');
       });
   });
} else {
 console.log('No students file found.');
}
});






// Call the password check before any commands

askForPassword().then(() => {


program.parse(process.argv);

})









// Function to convert JSON to Excel
function convertJsonToExcel(filePath, excelFilePath) {
try {
// Check if the JSON file exists
if (!fs.existsSync(filePath)) {
 console.error(`Error: The JSON file at ${filePath} does not exist.`);
 return;
}

// Read and parse the JSON file
const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));

// Check if the data is an array and contains data
if (!Array.isArray(jsonData) || jsonData.length === 0) {
 console.error('Error: The JSON file is empty or not an array.');
 return;
}

// Create a new workbook and worksheet
const workbook = xlsx.utils.book_new();
const worksheet = xlsx.utils.json_to_sheet(jsonData);

// Append the worksheet to the workbook
xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Write the workbook to an Excel file
xlsx.writeFile(workbook, excelFilePath);

console.log(`Excel file created at: ${excelFilePath}`);
} catch (error) {
console.error('Error:', error.message);
}
}








// Call the function
convertJsonToExcel(filePath, excelFilePath);





// Function to convert JSON to Excel
function convertJsonToExcel2(studentsFilePath, excelFilePathStudent) {
try {
// Check if the JSON file exists
if (!fs.existsSync(studentsFilePath)) {
 console.error(`Error: The JSON file at ${excelFilePathStudent} does not exist.`);
 return;
}

// Read and parse the JSON file
const jsonData = JSON.parse(fs.readFileSync(studentsFilePath, 'utf8'));

// Check if the data is an array and contains data
if (!Array.isArray(jsonData) || jsonData.length === 0) {
 console.error('Error: The JSON file is empty or not an array.');
 return;
}






// Create a new workbook and worksheet
const workbook = xlsx.utils.book_new();
const worksheet = xlsx.utils.json_to_sheet(jsonData);

// Append the worksheet to the workbook
xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Write the workbook to an Excel file
xlsx.writeFile(workbook, excelFilePathStudent);

console.log(`Excel file created at: ${excelFilePathStudent}`);
} catch (error) {
console.error('Error:', error.message);
}
}






// Call the function
convertJsonToExcel2(studentsFilePath, excelFilePathStudent);


console.log()
