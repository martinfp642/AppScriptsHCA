// Global variables for spreadsheets
const RESERVATIONS_SPREADSHEET_ID = '1QQEuUH3H9SmIpSahO3H7O1mcHzzVcfHA3TQl9-1hkvo';
const USERS_SPREADSHEET_ID = '1aLu5W8eq6i4CIOvm3F6D19n5xlUShN2T3HikLibQB3w';

// Definición de las franjas horarias
const TIME_SLOTS = [
  { label: "1era", ini: "07:45", fin: "08:25" },
  { label: "2da", ini: "08:25", fin: "09:05" },
  { label: "3era", ini: "09:10", fin: "09:50" },
  { label: "4ta", ini: "09:50", fin: "10:30" },
  { label: "5ta", ini: "10:35", fin: "11:15" },
  { label: "6ta", ini: "11:15", fin: "11:55" },
  { label: "7ma", ini: "12:15", fin: "12:55" },
  { label: "8va", ini: "12:55", fin: "13:35" },
  { label: "9na", ini: "13:40", fin: "14:20" },
  { label: "10ma", ini: "14:20", fin: "15:00" }
];

// Función para obtener la franja horaria actual
function getCurrentTimeSlot(date) {
  const now = date || new Date();
  const timeString = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
  
  for (let i = 0; i < TIME_SLOTS.length; i++) {
    const slot = TIME_SLOTS[i];
    if (timeString >= slot.ini && timeString <= slot.fin) {
      return slot.label;
    }
  }
  
  // Si está fuera de todas las franjas
  return "Fuera de horario";
}

// Función para calcular la duración en horas y minutos
function calculateDuration(startDate, endDate) {
  const diff = endDate.getTime() - startDate.getTime();
  const hours = Math.floor(diff / (1000 * 60 * 60));
  const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
  
  return {
    hours: hours,
    minutes: minutes,
    totalHours: hours + (minutes / 60),
    text: `${hours}h ${minutes}m`
  };
}

// Setup spreadsheets and required sheets
function setupSpreadsheets() {
  try {
    // First check if we can access both spreadsheets
    let reservationsSS, usersSS;
    
    try {
      reservationsSS = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
      Logger.log('Acceso exitoso a la hoja de reservas');
    } catch (e) {
      Logger.log('Error al acceder a la hoja de reservas: ' + e.toString());
      throw new Error('No se puede acceder a la hoja de reservas. Verifique que el ID sea correcto y que tenga permisos.');
    }
    
    try {
      usersSS = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
      Logger.log('Acceso exitoso a la hoja de usuarios');
    } catch (e) {
      Logger.log('Error al acceder a la hoja de usuarios: ' + e.toString());
      throw new Error('No se puede acceder a la hoja de usuarios. Verifique que el ID sea correcto y que tenga permisos.');
    }
    
    // Setup reservations spreadsheet
    // Check and create "Reservas alargues" sheet if it doesn't exist
    let reservationsSheet = reservationsSS.getSheetByName("Reservas alargues");
    if (!reservationsSheet) {
      Logger.log('Creando hoja de Reservas alargues');
      reservationsSheet = reservationsSS.insertSheet("Reservas alargues");
      // Set up headers - Agregando columnas para franjas horarias y duración
      reservationsSheet.getRange(1, 1, 1, 9).setValues([["Alargue", "Tipo Usuario", "Curso", "Solicitante", "Hora Devolución", "Hora Préstamo", "Docente", "Franja Préstamo", "Franja Devolución"]]);
      reservationsSheet.setFrozenRows(1);
    } else {
      // Verificamos si necesitamos actualizar los encabezados para añadir las nuevas columnas
      const headers = reservationsSheet.getRange(1, 1, 1, 9).getValues()[0];
      if (headers.length < 9 || headers[7] !== "Franja Préstamo" || headers[8] !== "Franja Devolución") {
        // Añadimos las nuevas columnas si no existen
        if (headers.length < 8) {
          reservationsSheet.getRange(1, 8).setValue("Franja Préstamo");
        }
        if (headers.length < 9) {
          reservationsSheet.getRange(1, 9).setValue("Franja Devolución");
        }
      }
    }
    
    // Check and create "lista negra" sheet if it doesn't exist
    let blacklistSheet = reservationsSS.getSheetByName("lista negra");
    if (!blacklistSheet) {
      Logger.log('Creando hoja de lista negra');
      blacklistSheet = reservationsSS.insertSheet("lista negra");
      // Set up headers
      blacklistSheet.getRange(1, 1, 1, 4).setValues([["Estudiante", "Fecha Sanción", "Reportado por", "Motivo"]]);
      blacklistSheet.setFrozenRows(1);
    }
    
    // Check and create "configuracion" sheet if it doesn't exist
    let configSheet = reservationsSS.getSheetByName("configuracion");
    if (!configSheet) {
      Logger.log('Creando hoja de configuración');
      configSheet = reservationsSS.insertSheet("configuracion");
      // Set up headers and default values for 6 extension cords
      configSheet.getRange(1, 1, 1, 2).setValues([["Parámetro", "Valor"]]);
      configSheet.getRange(2, 1, 6, 2).setValues([
        ["Alargue 1", "true"],
        ["Alargue 2", "true"],
        ["Alargue 3", "true"],
        ["Alargue 4", "true"],
        ["Alargue 5", "true"],
        ["Alargue 6", "true"]
      ]);
      configSheet.setFrozenRows(1);
    }
    
    // Check user spreadsheet sheets
    // We just check if they exist, we don't create them as they should have the student/teacher data
    const studentsSheet = usersSS.getSheetByName("estudiantes");
    if (!studentsSheet) {
      Logger.log('ADVERTENCIA: No se encontró la hoja de estudiantes');
    }
    
    const teachersSheet = usersSS.getSheetByName("docentes");
    if (!teachersSheet) {
      Logger.log('ADVERTENCIA: No se encontró la hoja de docentes');
    }
    
    Logger.log('Configuración de hojas completada exitosamente');
    return true;
  } catch (error) {
    Logger.log('Error en setupSpreadsheets: ' + error.toString());
    throw error; // Re-throw to be caught by the caller
  }
}

// Open the web app
function doGet() {
  try {
    // Asegurarse de que las hojas existan
    setupSpreadsheets();
    
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Sistema de Préstamo de Alargues')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('Error en doGet: ' + error.toString());
    // Crear una página de error simple para mostrar al usuario
    let output = HtmlService.createHtmlOutput(
      '<html><head><title>Error</title></head><body>' +
      '<h1>Error al iniciar la aplicación</h1>' +
      '<p>Ha ocurrido un error al acceder a las hojas de cálculo: ' + error.message + '</p>' +
      '<p>Por favor, verifique que los IDs de las hojas sean correctos y que tenga los permisos necesarios.</p>' +
      '<button onclick="window.location.reload()">Reintentar</button>' +
      '</body></html>'
    );
    output.setTitle('Error - Sistema de Préstamo de Alargues');
    return output;
  }
}

// Function to get enabled extension cords configuration
function getExtensionCordsConfig() {
  try {
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const configSheet = ss.getSheetByName("configuracion");
    
    if (!configSheet) {
      // If sheet doesn't exist, assume all cords are enabled
      return [1, 2, 3, 4, 5, 6].map(i => ({
        number: i,
        enabled: true
      }));
    }
    
    const data = configSheet.getRange(2, 1, 6, 2).getValues();
    const config = [];
    
    for (let i = 0; i < data.length; i++) {
      const cordName = data[i][0];
      const isEnabled = data[i][1].toString().toLowerCase() === 'true';
      const match = cordName.match(/Alargue (\d+)/);
      
      if (match) {
        const cordNumber = parseInt(match[1]);
        config.push({
          number: cordNumber,
          enabled: isEnabled
        });
      }
    }
    
    // Sort by cord number
    config.sort((a, b) => a.number - b.number);
    return config;
  } catch (error) {
    Logger.log('Error en getExtensionCordsConfig: ' + error.toString());
    // Default to all cords enabled if there's an error
    return [1, 2, 3, 4, 5, 6].map(i => ({
      number: i,
      enabled: true
    }));
  }
}

// Function to update extension cord config
function updateExtensionCordConfig(cordNumber, enabled) {
  try {
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    let configSheet = ss.getSheetByName("configuracion");
    
    if (!configSheet) {
      // Create config sheet if it doesn't exist
      configSheet = ss.insertSheet("configuracion");
      configSheet.getRange(1, 1, 1, 2).setValues([["Parámetro", "Valor"]]);
      
      // Set default values for all cords
      for (let i = 1; i <= 6; i++) {
        configSheet.getRange(i + 1, 1, 1, 2).setValues([["Alargue " + i, "true"]]);
      }
    }
    
    // Find the row for this cord
    const data = configSheet.getRange(2, 1, 6, 1).getValues();
    let rowFound = false;
    
    for (let i = 0; i < data.length; i++) {
      const match = data[i][0].match(/Alargue (\d+)/);
      if (match && parseInt(match[1]) === cordNumber) {
        configSheet.getRange(i + 2, 2).setValue(enabled.toString());
        rowFound = true;
        break;
      }
    }
    
    // If cord not found in config, add it
    if (!rowFound) {
      const lastRow = configSheet.getLastRow();
      configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([["Alargue " + cordNumber, enabled.toString()]]);
    }
    
    return {
      success: true,
      message: 'Configuración de alargue actualizada exitosamente.'
    };
  } catch (error) {
    Logger.log('Error en updateExtensionCordConfig: ' + error.toString());
    return {
      success: false,
      message: 'Error al actualizar la configuración: ' + error.toString()
    };
  }
}

// Function to get all courses from the students spreadsheet
function getCourses() {
  try {
    const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("estudiantes");
    
    if (!sheet) {
      // If the sheet doesn't exist, return an empty array
      Logger.log('La hoja "estudiantes" no existe');
      return [];
    }
    
    // Check if there are any rows in the sheet
    if (sheet.getLastRow() < 2) {
      return [];
    }
    
    const courses = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    // Remove duplicates and empty values
    const uniqueCourses = [...new Set(courses.flat().filter(Boolean))];
    return uniqueCourses;
  } catch (error) {
    Logger.log('Error en getCourses: ' + error.toString());
    return [];
  }
}

// Function to get students by course
function getStudentsByCourse(course) {
  try {
    const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("estudiantes");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // Filter students by course
    const students = data.filter(row => row[0] === course).map(row => row[1]);
    return students;
  } catch (error) {
    Logger.log('Error en getStudentsByCourse: ' + error.toString());
    return [];
  }
}

// Function to get all teachers - modificada para devolver una lista simple de docentes
function getTeachers() {
  try {
    const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("docentes");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    // Solo tomamos la columna B (nombres de docentes) ignorando la columna A (curso)
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    
    // Convertir a array simple y eliminar duplicados y valores vacíos
    const teachers = [...new Set(data.flat().filter(Boolean))];
    return teachers.sort(); // Devolvemos la lista ordenada alfabéticamente
    
  } catch (error) {
    Logger.log('Error en getTeachers: ' + error.toString());
    return [];
  }
}

// Function to get available extension cords
function getAvailableExtensionCords() {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    // First get configuration of enabled cords
    const cordsConfig = getExtensionCordsConfig();
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reservas alargues");
    
    // Initialize cords based on configuration
    const extensionCords = cordsConfig.map(config => ({
      number: config.number,
      available: config.enabled // Initial state based on enabled status
    }));
    
    // If no cords are configured or enabled, return empty array
    if (extensionCords.length === 0) {
      return [];
    }
    
    // If the sheet exists and has data, check which cords are currently on loan
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
      
      // Check which cords are currently on loan
      data.forEach(row => {
        if (row[0] && !row[4]) { // If there's an extension cord number and no return time
          const cordNumber = parseInt(row[0]);
          // Find this cord in our array
          const cordIndex = extensionCords.findIndex(cord => cord.number === cordNumber);
          if (cordIndex !== -1 && extensionCords[cordIndex].available) {
            extensionCords[cordIndex].available = false;
          }
        }
      });
    }
    
    // Filter to only return enabled cords
    return extensionCords.filter(cord => cordsConfig.find(c => c.number === cord.number)?.enabled);
  } catch (error) {
    Logger.log('Error en getAvailableExtensionCords: ' + error.toString());
    // Return empty array in case of error
    return [];
  }
}

// Function to check if a student is blacklisted
function isBlacklisted(studentName) {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const blacklistSheet = ss.getSheetByName("lista negra");
    
    if (!blacklistSheet || blacklistSheet.getLastRow() < 2) {
      return false;
    }
    
    const blacklist = blacklistSheet.getRange(2, 1, blacklistSheet.getLastRow() - 1, 1).getValues();
    
    return blacklist.flat().includes(studentName);
  } catch (error) {
    Logger.log('Error en isBlacklisted: ' + error.toString());
    return false;
  }
}

// Function to loan an extension cord
function loanExtensionCord(cordNumber, userType, course, studentName, teacherName) {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    // Check if the student is blacklisted
    if (userType === 'student' && isBlacklisted(studentName)) {
      return {
        success: false,
        message: 'Este estudiante está sancionado y no puede solicitar préstamos.'
      };
    }
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reservas alargues");
    
    // Record the loan
    const now = new Date();
    const timeSlot = getCurrentTimeSlot(now);
    
    const loanData = [
      cordNumber,
      userType,
      course,
      userType === 'student' ? studentName : teacherName,
      '', // Return time (empty initially)
      now, // Loan time
      teacherName, // Teacher name (if applicable)
      timeSlot,  // Franja horaria de préstamo
      ''  // Franja horaria de devolución (vacía inicialmente)
    ];
    
    // Add to the next empty row
    sheet.appendRow(loanData);
    
    return {
      success: true,
      message: 'Préstamo registrado correctamente.'
    };
  } catch (error) {
    Logger.log('Error en loanExtensionCord: ' + error.toString());
    return {
      success: false,
      message: 'Error al registrar préstamo: ' + error.toString()
    };
  }
}

// Function to return an extension cord
function returnExtensionCord(cordNumber, returnerName, wasBorrowedByReturner) {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reservas alargues");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: false,
        message: 'No se encontró un préstamo activo para este alargue.'
      };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    
    // Find the most recent loan for this cord
    let rowIndex = -1;
    let borrowerName = '';
    let loanTime = null;
    let loanTimeSlot = '';
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] == cordNumber && !data[i][4]) { // Found unreturned loan
        rowIndex = i + 2; // +2 because we start at row 2 and i is 0-based
        borrowerName = data[i][3];
        loanTime = new Date(data[i][5]);
        loanTimeSlot = data[i][7] || '';
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {
        success: false,
        message: 'No se encontró un préstamo activo para este alargue.'
      };
    }
    
    // Record the return time
    const now = new Date();
    const returnTimeSlot = getCurrentTimeSlot(now);
    
    // Calcular duración
    const duration = calculateDuration(loanTime, now);
    
    // Actualizar la hoja con la hora de devolución y la franja horaria
    sheet.getRange(rowIndex, 5).setValue(now); // Hora de devolución
    sheet.getRange(rowIndex, 9).setValue(returnTimeSlot); // Franja de devolución
    
    // Mensaje para mostrar en la interfaz
    let durationMessage = `Duración del préstamo: ${duration.text}`;
    let timeSlotMessage = `Franjas: ${loanTimeSlot} → ${returnTimeSlot}`;
    
    if (returnTimeSlot === "Fuera de horario") {
      timeSlotMessage = `Franjas: ${loanTimeSlot} → FUERA DEL HORARIO`;
    }
    
    // If the returner is not the borrower, add to blacklist
    if (!wasBorrowedByReturner) {
      const blacklistSheet = ss.getSheetByName("lista negra");
      blacklistSheet.appendRow([borrowerName, now, returnerName, 'No devolvió personalmente']);
      
      return {
        success: true,
        message: 'Alargue devuelto. El estudiante ' + borrowerName + ' ha sido sancionado por no devolverlo personalmente.\n\n' + durationMessage + '\n' + timeSlotMessage,
        duration: duration,
        loanTimeSlot: loanTimeSlot,
        returnTimeSlot: returnTimeSlot
      };
    }
    
    return {
      success: true,
      message: 'Alargue devuelto correctamente.\n\n' + durationMessage + '\n' + timeSlotMessage,
      duration: duration,
      loanTimeSlot: loanTimeSlot,
      returnTimeSlot: returnTimeSlot
    };
  } catch (error) {
    Logger.log('Error en returnExtensionCord: ' + error.toString());
    return {
      success: false,
      message: 'Error al registrar devolución: ' + error.toString()
    };
  }
}

// Admin functions

// Function to verify admin credentials
function verifyAdmin(username, password) {
  return username === 'admin' && password === 'admin';
}

// Function to get all current loans - VERSIÓN CORREGIDA
function getCurrentLoans() {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reservas alargues");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    
    // Filter to only show unreturned loans
    const currentLoans = data.filter(row => row[0] && !row[4]).map(row => {
      try {
        const loanTime = new Date(row[5]);
        const loanTimeStr = Utilities.formatDate(loanTime, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        
        return {
          cordNumber: parseInt(row[0]), // Convertir a entero
          userType: row[1],
          course: row[2],
          borrowerName: row[3],
          loanTime: loanTime.toLocaleString(),
          loanDate: loanTimeStr, // Fecha formateada para que funcione bien en el cliente
          teacherName: row[6] || 'N/A',
          timeSlot: row[7] || 'N/A'
        };
      } catch (err) {
        Logger.log('Error al procesar fila: ' + err.toString());
        // Si hay error, devolvemos un objeto con datos mínimos
        return null;
      }
    }).filter(item => item !== null); // Filtrar cualquier ítem nulo
    
    return currentLoans;
  } catch (error) {
    Logger.log('Error en getCurrentLoans: ' + error.toString());
    return [];
  }
}

// Function to get loan history by group
function getLoanHistoryByGroup(course) {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Reservas alargues");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    
    // Filter by course
    const loanHistory = data
      .filter(row => row[0] && row[2] === course)
      .map(row => {
        try {
          const loanTime = new Date(row[5]);
          const returnTime = row[4] ? new Date(row[4]) : null;
          let durationText = '';
          
          if (returnTime) {
            const duration = calculateDuration(loanTime, returnTime);
            durationText = duration.text;
          }
          
          return {
            cordNumber: parseInt(row[0]),
            userType: row[1],
            borrowerName: row[3],
            loanTime: loanTime.toLocaleString(),
            returnTime: returnTime ? returnTime.toLocaleString() : 'No devuelto',
            teacherName: row[6] || 'N/A',
            loanTimeSlot: row[7] || 'N/A',
            returnTimeSlot: row[8] || 'N/A',
            duration: durationText,
            isReturned: !!row[4]
          };
        } catch (err) {
          Logger.log('Error al procesar fila en historial: ' + err.toString());
          return null;
        }
      }).filter(item => item !== null); // Filtrar cualquier ítem nulo
    
    return loanHistory;
  } catch (error) {
    Logger.log('Error en getLoanHistoryByGroup: ' + error.toString());
    return [];
  }
}

// Function to get blacklisted students
function getBlacklistedStudents() {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("lista negra");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    // Map the blacklist data
    const blacklist = data.filter(row => row[0]).map(row => {
      try {
        const sanctionDate = new Date(row[1]);
        return {
          studentName: row[0],
          sanctionDate: sanctionDate.toLocaleString(),
          reportedBy: row[2] || 'N/A',
          reason: row[3] || 'N/A'
        };
      } catch (err) {
        Logger.log('Error al procesar sanción: ' + err.toString());
        return null;
      }
    }).filter(item => item !== null); // Filtrar cualquier ítem nulo
    
    return blacklist;
  } catch (error) {
    Logger.log('Error en getBlacklistedStudents: ' + error.toString());
    return [];
  }
}

// Function to remove a sanction
function removeSanction(studentName) {
  try {
    setupSpreadsheets(); // Ensure sheets exist
    
    const ss = SpreadsheetApp.openById(RESERVATIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("lista negra");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: false,
        message: 'No se encontró una sanción para ' + studentName
      };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    // Find the row with this student
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === studentName) {
        // Delete the row
        sheet.deleteRow(i + 2); // +2 because we start at row 2 and i is 0-based
        return {
          success: true,
          message: 'Sanción removida para ' + studentName
        };
      }
    }
    
    return {
      success: false,
      message: 'No se encontró una sanción para ' + studentName
    };
  } catch (error) {
    Logger.log('Error en removeSanction: ' + error.toString());
    return {
      success: false,
      message: 'Error al remover sanción: ' + error.toString()
    };
  }
}
