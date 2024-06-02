import * as readline from 'readline';
import { MongoClient } from 'mongodb';
import * as ExcelJS from 'exceljs';
import { spawn } from 'child_process';


interface Para {
    time: string;
    group: string;
    classroom: string;
    subject: string;
    teacher: string;
}
// Описывает расписание на один день, может содержать несколько записей Para.
interface Day {
    entries: Para[];
}

interface Week {
  [day: string]: Day[];
}

interface Weekly {
    oddWeek: Week;
    evenWeek: Week;
}

interface Teacher {
  name: string;
  schedule: Weekly;
}

interface GroupSchedule {
  group: string;
  schedule: Weekly;
}

function createSchedule(): Weekly {
  return {
    oddWeek: {
      Понедельник: [],
      Вторник: [],
      Среда: [],
      Четверг: [],
      Пятница: [],
      Суббота: []
    },
    evenWeek: {
      Понедельник: [],
      Вторник: [],
      Среда: [],
      Четверг: [],
      Пятница: [],
      Суббота: []
    }
  };
}

const daysOfWeek: (keyof Week)[] = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
const times = ['08.00-09.35', '09.45-11.20', '11.30-13.05', '13.55-15.30', '15.40-17.15'];

// Функция для чтения Excel файла и преобразования данных в структурированный формат
async function readerExcel(filePath: string): Promise<Teacher[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Преподаватели');
  const teachers: Teacher[] = [];

  if (!worksheet) {
    throw new Error("Лист 'Преподаватели' не найден.");
  }

  let currentRow = 7; // Начинаем с ячейки B7
  while (worksheet.getCell(`B${currentRow}`).value) {
    const name = worksheet.getCell(`B${currentRow}`).value as string;
    const schedule = createSchedule();

    for (let dayIndex = 0; dayIndex < daysOfWeek.length; dayIndex++) {
      const dayName = daysOfWeek[dayIndex];
      for (let timeIndex = 0; timeIndex < times.length; timeIndex++) {
        const time = times[timeIndex];
        const row = currentRow + 2 + timeIndex * 4; // Смещение на 4 строки для каждого временного слота
        const col = 2 + dayIndex; // Смещение на 2 столбца для каждого дня недели

        const oddWeekCellContent = worksheet.getCell(row, col).value as string || '';
        const oddWeekSubject = worksheet.getCell(row + 1, col).value as string || '';
        const evenWeekCellContent = worksheet.getCell(row + 2, col).value as string || '';
        const evenWeekSubject = worksheet.getCell(row + 3, col).value as string || '';

        // Функция для разделения строки на группу и аудиторию
        const splitGroupAndroom = (content: string) => {
          const parts = content.split(' а.');
          return {
            group: parts[0].trim(),
            classroom: parts.length > 1 ? 'а.' + parts[1].trim() : '',
          };
        };

        const oddWeekParaData = splitGroupAndroom(oddWeekCellContent);
        const evenWeekParaData = splitGroupAndroom(evenWeekCellContent);

        const oddWeekPara: Para = { time, ...oddWeekParaData, subject: oddWeekSubject, teacher: name };
        const evenWeekPara: Para = { time, ...evenWeekParaData, subject: evenWeekSubject, teacher: name };

        schedule.oddWeek[dayName].push({ entries: [oddWeekPara] });
        schedule.evenWeek[dayName].push({ entries: [evenWeekPara] });
      }
    }

    teachers.push({ name, schedule });
    currentRow += 32; // Переход к следующему преподавателю
  }

  return teachers;
}


async function insertMongoDB(filePath: string): Promise<void> {
  const newTeacherSchedules = await readerExcel(filePath); 

  const url = "mongodb://root:example@localhost:27017/";
  const dbName = 'Curs';
  const collectionNameT = 'Teachers';

  const client = new MongoClient(url);
  try {
    await client.connect();
    console.log("Успешное подключение к MongoDB");
    const db = client.db(dbName);
    const collection = db.collection<Teacher>(collectionNameT);

    for (const newTeacherSchedule of newTeacherSchedules) {
      const existingTeacherSchedule = await collection.findOne({ name: newTeacherSchedule.name });

      if (existingTeacherSchedule) {
        // Объединяем расписание для нечетной и четной недели
        for (const weekType of ['oddWeek', 'evenWeek'] as const) {
          for (const day of daysOfWeek) {
            const existingDaySchedules = existingTeacherSchedule.schedule[weekType][day];
            const newDaySchedules = newTeacherSchedule.schedule[weekType][day];
  
            // Обходим все записи для данного дня
            newDaySchedules.forEach((newDaySchedule) => {
              newDaySchedule.entries.forEach((newEntry) => {
                // Ищем соответствующий Day и entry в существующем расписании
                const existingDaySchedule = existingDaySchedules.find(ed => ed.entries.some(e => e.time === newEntry.time));
                if (!existingDaySchedule) {
                  // Если нет соответствующего Day, добавляем весь Day
                  existingDaySchedules.push(newDaySchedule);
                } else {
                  // Если Day найден, ищем entry для обновления
                  const existingEntryIndex = existingDaySchedule.entries.findIndex(e => e.time === newEntry.time);
                  if (existingEntryIndex !== -1) {
                    // Обновляем существующую запись, если новая информация более полная
                    const existingEntry = existingDaySchedule.entries[existingEntryIndex];
                    if (newEntry.group && (!existingEntry.group || newEntry.group.length > existingEntry.group.length)) {
                      existingEntry.group = newEntry.group;
                    }
                    if (newEntry.classroom && (!existingEntry.classroom || newEntry.classroom.length > existingEntry.classroom.length)) {
                      existingEntry.classroom = newEntry.classroom;
                    }
                    if (newEntry.subject.trim() !== '' && (!existingEntry.subject || newEntry.subject.length > existingEntry.subject.length)) {
                      existingEntry.subject = newEntry.subject;
                    }
                    // Обновляем teacher только если поле не пустое
                    if (newEntry.teacher.trim() !== '') {
                      existingEntry.teacher = newEntry.teacher;
                    }
                  }
                }
              });
            });
          }
        }

        // Обновляем расписание преподавателя в базе данных
        await collection.updateOne(
          { name: newTeacherSchedule.name },
          { $set: { schedule: existingTeacherSchedule.schedule } }
        );
      } else {
        // Если расписания в базе данных нет, то используем новое расписание
        await collection.insertOne(newTeacherSchedule);
      }
    }
    console.log("Данные успешно вставлены в MongoDB");
  } catch (err) {
    console.error("Произошла ошибка при вставке или обновлении данных в MongoDB:", err);
  } finally {
    await client.close();
  }
}


async function findTeacherSchedule(client: MongoClient, dbName: string, collectionName: string, teacherIdentifier: string): Promise<Teacher | null> {
  const db = client.db(dbName);
  const collection = db.collection(collectionName);

  // Используем регулярное выражение для поиска по фамилии или инициалам
  const regex = new RegExp(teacherIdentifier, 'i');
  const teacher = await collection.findOne({ name: regex }) as unknown as Teacher | null;

  return teacher;
}


async function findGroupSchedule(client: MongoClient, dbName: string, collectionName: string, groupIdentifier: string): Promise<GroupSchedule | null> {
  const db = client.db(dbName);
  const collection = db.collection(collectionName);

  const teachers = (await collection.find({}).toArray()) as unknown as Teacher[];
  let scheduleFound = false;
  const groupSchedule: GroupSchedule = {
    group: groupIdentifier,
    schedule: createSchedule()
  };

  // Функция для проверки, содержит ли строка идентификатор группы
  const containsGroup = (groupString: string, groupIdentifier: string): boolean => {
    const groupParts = groupString.split(/[~;]/); 
    return groupParts.includes(groupIdentifier);
  };

  teachers.forEach(teacher => {
    Object.keys(teacher.schedule.oddWeek).forEach(day => {
      teacher.schedule.oddWeek[day].forEach(daySchedule => {
        daySchedule.entries.forEach(entry => {
          if (containsGroup(entry.group, groupIdentifier)) {
            groupSchedule.schedule.oddWeek[day].push({ entries: [entry] });
            scheduleFound = true;
          }
    })})})

    Object.keys(teacher.schedule.evenWeek).forEach(day => {
      teacher.schedule.evenWeek[day].forEach(daySchedule => {
        daySchedule.entries.forEach(entry => {
          if (containsGroup(entry.group, groupIdentifier)) {
            groupSchedule.schedule.evenWeek[day].push({ entries: [entry] });
            scheduleFound = true;
          }
  })})})})

  return scheduleFound ? groupSchedule : null;
}


type ScheduleType = Teacher | GroupSchedule;

async function createExcelFile(schedule: ScheduleType, type: 'teacher' | 'group'): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Расписание');

  worksheet.columns = [
    { key: 'A', width: 19 },
    { key: 'B', width: 24 },
    { key: 'C', width: 24 },
    { key: 'D', width: 24 },
    { key: 'E', width: 24 },
    { key: 'F', width: 24 },
    { key: 'G', width: 24 }
  ];

  // Записываем имя преподавателя в ячейку A1
  worksheet.getCell('A1').value = type === 'teacher' ? (schedule as Teacher).name : (schedule as GroupSchedule).group;

  // Записываем дни недели начиная с ячейки B2
  let currentColumn = 'B';
  for (const day of daysOfWeek) {
    worksheet.getCell(`${currentColumn}2`).value = day;
    currentColumn = String.fromCharCode(currentColumn.charCodeAt(0) + 1); // Переходим к следующему столбцу
  }

  // Записываем времена пар начиная с ячейки A3
  currentColumn = 'B';
  for (const day of daysOfWeek) {
    let currentTimeRow = 3;
    for (const time of times) {
      worksheet.getCell(`A${currentTimeRow}`).value = time;
      // Получаем данные для нечетной и четной недели
      let oddWeekDay, evenWeekDay, oddWeekPara, evenWeekPara;
      if (type === 'teacher') {
        oddWeekDay = (schedule as Teacher).schedule.oddWeek[day];
        evenWeekDay = (schedule as Teacher).schedule.evenWeek[day];
        oddWeekPara = oddWeekDay.flatMap(d => d.entries).find((entry: Para) => entry.time === time);
        evenWeekPara = evenWeekDay.flatMap(d => d.entries).find((entry: Para) => entry.time === time);
      } else {
        oddWeekDay = (schedule as GroupSchedule).schedule.oddWeek[day];
        evenWeekDay = (schedule as GroupSchedule).schedule.evenWeek[day];
        oddWeekPara = oddWeekDay.flatMap(d => d.entries).find((entry: Para) => entry.time === time);
        evenWeekPara = evenWeekDay.flatMap(d => d.entries).find((entry: Para) => entry.time === time);
      }

      // Записываем данные для нечетной и четной недели
      worksheet.getCell(`${currentColumn}${currentTimeRow}`).value = oddWeekPara ? `${type === 'teacher' ? oddWeekPara.group : oddWeekPara.teacher} ${oddWeekPara.classroom}` : '';
      worksheet.getCell(`${currentColumn}${currentTimeRow + 1}`).value = oddWeekPara?.subject || '';
      worksheet.getCell(`${currentColumn}${currentTimeRow + 2}`).value = evenWeekPara ? `${type === 'teacher' ? evenWeekPara.group : evenWeekPara.teacher} ${evenWeekPara.classroom}` : '';
      worksheet.getCell(`${currentColumn}${currentTimeRow + 3}`).value = evenWeekPara?.subject || '';

      currentTimeRow += 4; // Переходим к следующему временному слоту
    }
    currentColumn = String.fromCharCode(currentColumn.charCodeAt(0) + 1); // Переходим к следующему дню недели
  }

  // Устанавливаем жирные границы для заголовков дней недели
  for (let col = 1; col <= daysOfWeek.length + 1; col++) {
    const cell = worksheet.getCell(2, col);
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
  }

  for (let row = 3; row <= 22; row += 4) {
    for (let col = 1; col <= daysOfWeek.length + 1; col++) {
      // Устанавливаем верхнюю и правую границы для первой ячейки каждого временного блока
      const firstCellInBlock = worksheet.getCell(row, col);
      firstCellInBlock.border = {
        top: { style: 'medium' },
        right: { style: 'medium' }
      };

      // Устанавливаем только правую границу для следующих трех ячеек в каждом временном блоке
      for (let i = 1; i <= 3; i++) {
        const cellWithRightBorder = worksheet.getCell(row + i, col);
        cellWithRightBorder.border = {
          right: { style: 'medium' }
        };
      }
    }
  }

  const lastRow = worksheet.getRow(22);
  lastRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.border = {
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
  });

  const fileName = `${type === 'teacher' ? (schedule as Teacher).name : (schedule as GroupSchedule).group}.xlsx`.replace(/[/\\?%*:|"<>]/g, '_').replace(/\.\.xlsx$/, '.xlsx');
  await workbook.xlsx.writeFile(fileName);

// Открытие файла с помощью программы по умолчанию
  const openFile = spawn('cmd.exe', ['/c', 'start', 'excel', `"${fileName}"`], { shell: true });

  openFile.on('error', (error: Error) => {
    console.error(`Не удалось открыть файл: ${error}`);
  });
}


async function main() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  const question = (query: string): Promise<string> => {
    return new Promise(resolve => {
      rl.question(query, resolve);
    });
  };

  const filePath = await question('Введите путь к файлу: ');
  await insertMongoDB(filePath);

  const identifier = await question('Введите фамилию преподавателя или номер группы: ');

  const url = "mongodb://root:example@localhost:27017/";
  const dbName = 'Curs';
  const collectionNameT = 'Teachers';

  const client = new MongoClient(url);
try {
  await client.connect();
  if (/^\d+[а-яА-Яa-zA-Z]+$/.test(identifier)) {
    // Идентификатор является номером группы
    const groupSchedule = await findGroupSchedule(client, dbName, collectionNameT, identifier);
      if (groupSchedule) {
      await createExcelFile(groupSchedule, 'group'); // Функция для группы
    } else {
      console.log("Расписание для группы не найдено.");
    }
  } else {
    // Идентификатор является фамилией преподавателя
    const teacher = await findTeacherSchedule(client, dbName, collectionNameT, identifier);
    if (teacher) {
      await createExcelFile(teacher, 'teacher'); // Функция для преподавателя
    } else {
      console.log("Преподаватель не найден.");
    }
  }
  } catch (err) {
    console.error("Произошла ошибка:", err);
  } finally {
    await client.close();
    rl.close();
  }
}

main().catch(console.error); 
