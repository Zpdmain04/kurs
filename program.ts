import * as fs from 'fs';
import * as yaml from 'js-yaml';
import { MongoClient } from 'mongodb';
import * as xlsx from 'xlsx';
import {Command} from 'commander';

class Student {
  name: string;
  disciplines: Discipline[]
  assignments: {
    discipline: Discipline
    assignment: string
    date: Date
    grade: number
  }[]
  birthday: Date
  faculty: string

  constructor(
    name: string,
    disciplines: Discipline[] = [],
    assignments: { discipline: Discipline; assignment: string; date: Date; grade: number; }[] = [],
    birthday: Date, 
    faculty: string
    ) {
    this.name = name;
    this.disciplines = disciplines
    this.assignments = assignments
    this.birthday = birthday
    this.faculty = faculty
  }
}

class Discipline {
  name: string
  assignments: {
    name: string
    deadline: Date
    maxGrade: number
  }[];

  constructor(
    name: string, 
    assignments: { name: string; deadline: Date; maxGrade: number; }[] = []
    ) {
    this.name = name
    this.assignments = assignments
  }
}

class MongoDB {
  client: MongoClient
  db: any

  constructor(
    server: string, 
    dbName: string
    ) {
    this.client = new MongoClient(server)
    this.db = this.client.db(dbName)
  }

  async connect() {
    await this.client.connect()
  }

  async disconnect() { 
    await this.client.close()
  }

  async saveData(collname: string, data: any) {
    const collection = this.db.collection(collname);
    for (const item of data) {
      const origdoc = await collection.findOne({name: item.name});
      if (!origdoc) {
        await collection.insertOne(item);
        continue
      }
      const updoc = { ...origdoc, ...item };
      if (JSON.stringify(origdoc) === JSON.stringify(updoc)) {
        continue
      }
      await collection.updateOne({name: item.name}, {$set: item});
    }
  }

  async getData(collname: string): Promise<any[]> {
    const collection = this.db.collection(collname)
    return await collection.find({}).toArray()
  }
  
}

function createxcel(data: any[], filename: string) {
  const wb = xlsx.utils.book_new()
  const ws = xlsx.utils.aoa_to_sheet(data)
  xlsx.utils.book_append_sheet(wb, ws, filename)
  xlsx.writeFile(wb, `${filename}.xlsx`)
}

async function loadStudents() {
  try {
    if (!fs.existsSync('students.yaml')) {
      throw new Error("Файл 'students.yaml' не найден.");
    }
    const MONGO = new MongoDB('mongodb://root:example@127.0.0.1:27017/', 'Students_n_Disciplines')
    await MONGO.connect()
    const studentsData = yaml.load(fs.readFileSync('students.yaml', 'utf-8')) as Student[]
    await MONGO.saveData('Students', studentsData)
    console.log('Студенты успешно загружены в MongoDB.')
    await MONGO.disconnect()
  } catch (e) {
    console.error(e)
  }
}

async function loadDisciplines() {
  try {
    if (!fs.existsSync('disciplines.yaml')) {
      throw new Error("Файл 'disciplines.yaml' не найден.");
    }
    const MONGO = new MongoDB('mongodb://root:example@127.0.0.1:27017/', 'Students_n_Disciplines')
    await MONGO.connect()
    const disciplinesData = yaml.load(fs.readFileSync('disciplines.yaml', 'utf-8')) as Discipline[]
    await MONGO.saveData('Disciplines', disciplinesData)
    console.log('Дисциплины успешно загружены в MongoDB.')
    await MONGO.disconnect()
  } catch (e) {
    console.error(e)
  }
}


async function exportStudents() {
  const MONGO = new MongoDB('mongodb://root:example@127.0.0.1:27017/', 'Students_n_Disciplines')
  await MONGO.connect()
  const students = await MONGO.getData('Students')
  const disciplines = await MONGO.getData('Disciplines')
  const studentDisciplinesData: any[][] = []
  for (const student of students) {
    const studentDisciplines = student.disciplines.map((disciplineName: any) => {
      const discipline = disciplines.find(d => d.name === disciplineName)
      return discipline ? discipline.name : null;
    })
    studentDisciplinesData.push([student.name, ...studentDisciplines])
  }
  createxcel(studentDisciplinesData, 'Students')
  console.log('Данные о студентах успешно выгружены в файл Students.xlsx.')
  await MONGO.disconnect()
}

async function exportDisciplines() {
  const MONGO = new MongoDB('mongodb://root:example@127.0.0.1:27017/', 'Students_n_Disciplines')
  await MONGO.connect()
  const disciplines = await MONGO.getData('Disciplines')
  const disciplineAssignments = disciplines.map(discipline => [
    discipline.name,
    ...discipline.assignments.map((assignment: { name: any; }) => assignment.name),
  ]);
  createxcel(disciplineAssignments, 'Disciplines')
  console.log('Данные о дисциплинах успешно выгружены в файл Disciplines.xlsx.')
  await MONGO.disconnect()
}
const program = new Command()


program
  .command('ls')
  .description('Загрузить данные студентов из YAML файла в MongoDB')
  .action(loadStudents)

program
  .command('ld')
  .description('Загрузить данные дисциплин из YAML файла в MongoDB')
  .action(loadDisciplines)

program
  .command('es')
  .description('Экспортировать данные о студентах в Excel файл')
  .action(exportStudents)

program
  .command('ed')
  .description('Экспортировать данные о дисциплинах в Excel файл')
  .action(exportDisciplines)

program.parse(process.argv)

