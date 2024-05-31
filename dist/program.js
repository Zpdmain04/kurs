"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __values = (this && this.__values) || function(o) {
    var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
    if (m) return m.call(o);
    if (o && typeof o.length === "number") return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
    throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
};
var __read = (this && this.__read) || function (o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var yaml = require("js-yaml");
var mongodb_1 = require("mongodb");
var xlsx = require("xlsx");
var Student = /** @class */ (function () {
    function Student(name, disciplines, assignments, birthday, faculty) {
        if (disciplines === void 0) { disciplines = []; }
        if (assignments === void 0) { assignments = []; }
        this.name = name;
        this.disciplines = disciplines;
        this.assignments = assignments;
        this.birthday = birthday;
        this.faculty = faculty;
    }
    return Student;
}());
var Discipline = /** @class */ (function () {
    function Discipline(name, assignments) {
        if (assignments === void 0) { assignments = []; }
        this.name = name;
        this.assignments = assignments;
    }
    return Discipline;
}());
var MongoDB = /** @class */ (function () {
    function MongoDB(server, dbName) {
        this.client = new mongodb_1.MongoClient(server);
        this.db = this.client.db(dbName);
    }
    MongoDB.prototype.connect = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.client.connect()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    MongoDB.prototype.disconnect = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.client.close()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    MongoDB.prototype.saveData = function (collname, data) {
        return __awaiter(this, void 0, void 0, function () {
            var collection, data_1, data_1_1, item, origdoc, updoc, e_1_1;
            var e_1, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        collection = this.db.collection(collname);
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 9, 10, 11]);
                        data_1 = __values(data), data_1_1 = data_1.next();
                        _b.label = 2;
                    case 2:
                        if (!!data_1_1.done) return [3 /*break*/, 8];
                        item = data_1_1.value;
                        return [4 /*yield*/, collection.findOne({ name: item.name })];
                    case 3:
                        origdoc = _b.sent();
                        if (!!origdoc) return [3 /*break*/, 5];
                        return [4 /*yield*/, collection.insertOne(item)];
                    case 4:
                        _b.sent();
                        _b.label = 5;
                    case 5:
                        updoc = __assign(__assign({}, origdoc), item);
                        if (JSON.stringify(origdoc) === JSON.stringify(updoc)) {
                            return [3 /*break*/, 7];
                        }
                        return [4 /*yield*/, collection.updateOne({ name: item.name }, { $set: item })];
                    case 6:
                        _b.sent();
                        _b.label = 7;
                    case 7:
                        data_1_1 = data_1.next();
                        return [3 /*break*/, 2];
                    case 8: return [3 /*break*/, 11];
                    case 9:
                        e_1_1 = _b.sent();
                        e_1 = { error: e_1_1 };
                        return [3 /*break*/, 11];
                    case 10:
                        try {
                            if (data_1_1 && !data_1_1.done && (_a = data_1.return)) _a.call(data_1);
                        }
                        finally { if (e_1) throw e_1.error; }
                        return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    };
    MongoDB.prototype.getData = function (collname) {
        return __awaiter(this, void 0, void 0, function () {
            var collection;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        collection = this.db.collection(collname);
                        return [4 /*yield*/, collection.find({}).toArray()];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    return MongoDB;
}());
function createxcel(data, filename) {
    var wb = xlsx.utils.book_new();
    var ws = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(wb, ws, filename);
    xlsx.writeFile(wb, "".concat(filename, ".xlsx"));
}
function main() {
    return __awaiter(this, void 0, void 0, function () {
        var MONGO, studentsData, disciplinesData, students, disciplines_1, studentDisciplinesData, students_1, students_1_1, student, studentDisciplines, disciplineAssignments, e_2;
        var e_3, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 7, , 8]);
                    MONGO = new MongoDB('mongodb://root:example@127.0.0.1:27017/', 'Students_n_Disciplines');
                    return [4 /*yield*/, MONGO.connect()];
                case 1:
                    _b.sent();
                    studentsData = yaml.load(fs.readFileSync('students.yaml', 'utf-8'));
                    return [4 /*yield*/, MONGO.saveData('Students', studentsData)];
                case 2:
                    _b.sent();
                    disciplinesData = yaml.load(fs.readFileSync('disciplines.yaml', 'utf-8'));
                    return [4 /*yield*/, MONGO.saveData('Disciplines', disciplinesData)];
                case 3:
                    _b.sent();
                    return [4 /*yield*/, MONGO.getData('Students')];
                case 4:
                    students = _b.sent();
                    return [4 /*yield*/, MONGO.getData('Disciplines')];
                case 5:
                    disciplines_1 = _b.sent();
                    studentDisciplinesData = [];
                    try {
                        for (students_1 = __values(students), students_1_1 = students_1.next(); !students_1_1.done; students_1_1 = students_1.next()) {
                            student = students_1_1.value;
                            studentDisciplines = student.disciplines.map(function (disciplineName) {
                                var discipline = disciplines_1.find(function (d) { return d.name === disciplineName; });
                                return discipline ? discipline.name : null;
                            });
                            studentDisciplinesData.push(__spreadArray([student.name], __read(studentDisciplines), false));
                        }
                    }
                    catch (e_3_1) { e_3 = { error: e_3_1 }; }
                    finally {
                        try {
                            if (students_1_1 && !students_1_1.done && (_a = students_1.return)) _a.call(students_1);
                        }
                        finally { if (e_3) throw e_3.error; }
                    }
                    createxcel(studentDisciplinesData, 'Students');
                    disciplineAssignments = disciplines_1.map(function (discipline) { return __spreadArray([
                        discipline.name
                    ], __read(discipline.assignments.map(function (assignment) { return assignment.name; })), false); });
                    createxcel(disciplineAssignments, 'Disciplines');
                    return [4 /*yield*/, MONGO.disconnect()];
                case 6:
                    _b.sent();
                    return [3 /*break*/, 8];
                case 7:
                    e_2 = _b.sent();
                    console.error(e_2);
                    return [3 /*break*/, 8];
                case 8: return [2 /*return*/];
            }
        });
    });
}
main();
//# sourceMappingURL=program.js.map