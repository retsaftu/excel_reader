const Excel = require('exceljs');
const MongoClient = require("mongodb").MongoClient;
const NumberInt = require("mongodb").Int32;

const  ObjectID = require('mongodb').ObjectId;


var wb = new Excel.Workbook();
var path = require('path');
var filePath = path.resolve(__dirname,'sample.xlsx');//nome of excel file
let arr = [];

   
const url = "mongodb://129.200.9.27:40508/";
const mongoClient = new MongoClient(url);
 
async function run(query) {
    try {
        await mongoClient.connect();
        const db = mongoClient.db("agisads");
        const collection = db.collection("dictionaryPassports");
        let str = 'Руслан';
        let str2 = 'Жакупов';
        const name = await  collection.find(query).toArray();
        // const name = await  collection.find({templateId: ObjectID("6093860155ddb8004538cd64"),$and : [{name: {'$regex': str}}, {name: {'$regex': str2}}]}).toArray();
        // const name = await  collection.find({templateId: ObjectID("6093860155ddb8004538cd64"),name: {'$regex': str}}).toArray();

        // console.log(name);
        return name;
         
    }catch(err) {
        console.log(err);
    } 
}
async function addData(arr) {
    try {
        await mongoClient.connect();
        const db = mongoClient.db("agisads");
        const collection = db.collection("dictionaryPassports");
        // create an array of documents to insert
      const docs = (JSON.parse(arr)).save();
      console.log(arr);
      const result = await collection.insertMany(docs);
      console.log(`${result.insertedCount} documents were inserted`);
    } finally {
      await mongoClient.close();
    }
  }
wb.xlsx.readFile(filePath).then(async function(){

    var sh = wb.getWorksheet("Worksheet");//name of list in excel

    // sh.getRow(1).getCell(2).value = 32;
    wb.xlsx.writeFile("sample2.xlsx");
    // console.log("Row-3 | Cell-2 - "+sh.getRow(1).getCell(2).value);

    // console.log("Row-3 | Cell-2 - "+sh.getRow(3).getCell(2).value);

    // console.log(sh.rowCount);
    //Get all the rows data [1st and 2nd column]
    for (i = 2; i <= 657; i++) {
        let application_number=sh.getRow(i).getCell(1).value;
        if(!application_number){
            application_number=null;
        }
        let author_name=sh.getRow(i).getCell(2).value+"";
        if(!author_name){
            author_name=null;
        }
        // let matching_string=sh.getRow(i).getCell(2).value+"";

        // console.log(author_name);
        // console.log(typeof(author_name));
        // console.log(author_name.indexOf('Синий'));
        let firstSpace = author_name.indexOf(' ');
        let firstName = author_name.substring(0, firstSpace);
        let lastName = firstName.replace(/\s/g, '');
        let secondSpace = author_name.indexOf(" ", firstSpace + 1);
        let middleName = firstName.replace(/\s/g, '');
        if (secondSpace < 0) { 
            lastName = author_name.substring(firstSpace);
            lastName = lastName.replace(/\s/g, '');
        }
        else {
            middleName = author_name.substring(firstSpace, secondSpace);
            lastName = author_name.substring(secondSpace);
            lastName = lastName.replace(/\s/g, '');
            firstName = firstName.replace(/\s/g, '');
            middleName = middleName.replace(/\s/g, '');
        }
        if(firstSpace==-1){
            firstName=author_name+'';
            lastName=author_name+'';
            middleName=author_name+'';
        }
        // console.log(lastName,firstName,middleName);
        let query = {templateId: ObjectID("6093860155ddb8004538cd64"),$and : [{name: {'$regex': firstName}}, {name: {'$regex': lastName}}, {name: {'$regex': middleName}}]};
        let author_id = await run(query);
        if(!author_id){
            author_id=null;
        }
        // console.log(author_id[0]._id);
        // const name = await  collection.find({templateId: ObjectID("6093860155ddb8004538cd64"),$and : [{name: {'$regex': str}}, {name: {'$regex': str2}}]}).toArray();

        let reconcile_arr=[];
        // for(i=0;i<)
        reconcile = sh.getRow(i).getCell(5).value;
        // console.log(reconcile.replace(/ *\([^)]*\) */g, ""));
        reconcile = reconcile.replace(/ *\([^)]*\) */g, "");
        for(let j=0;j<reconcile.split(",").length;j++){
            let fio = reconcile.split(', ');
            // fio[j]=fio[j].split(/\s*;\s*/,1);
            let reconcile_Name=fio[j].split(' ');
            // console.log(fio);
            // console.log(reconcile_Name);
            let query = {templateId: ObjectID("609248cb55ddb8004538cd63"),$and : [{name: {'$regex': reconcile_Name[0]}}, {name: {'$regex': reconcile_Name[1]}}]};
            let reconcile_id = await run(query);
            if(!reconcile_id){
                reconcile_id=null;
            }
            reconcile_arr[j]={
                "_id" : reconcile_id[0]._id,
                "passportType" : 2,
                "templateId" : "609248cb55ddb8004538cd63"
            }
            // console.log(reconcile_arr[j].passportType);
            // let space=reconcile.indexOf(' ');
            // let space_second=reconcile.indexOf(' ',space+1)
            // let reconcile_Surname = reconcile.substring(reconcile.indexOf(" ", space+j), reconcile.indexOf(" ", space_second+j));
            // // let reconcile_Name;
            // console.log(reconcile_Surname);
            // console.log(reconcile_Name);
        }

        let area = sh.getRow(i).getCell(8).value+"";
        let number_disconnect = sh.getRow(i).getCell(9).value;
        let mgd = sh.getRow(i).getCell(10).value;
        let chgd = sh.getRow(i).getCell(11).value;
        let detsad = sh.getRow(i).getCell(12).value;
        let hospital = sh.getRow(i).getCell(13).value;
        let school = sh.getRow(i).getCell(14).value;
        let other = sh.getRow(i).getCell(15).value;

        let executor = sh.getRow(i).getCell(17).value+"";
        console.log(executor);
        let executor_arr = [];
        for(let j=0;j<executor.split(",").length;j++){
            let fio = executor.split(', ');
            // fio[j]=fio[j].split(/\s*;\s*/,1);
            let reconcile_Name=fio[j].split(' ');
            // console.log(reconcile_Name);
            console.log(fio);
            // console.log("reconcile_Name",reconcile_Name);
            let query_executor = {templateId: ObjectID("6093869255ddb8004538cd66"),$and : [{name: {'$regex': fio[j]}}]};
            let executor_id = await run(query_executor);
            // console.log(executor_id);
            executor_arr[j] = {
                "_id" : executor_id[0]._id,
                "passportType" : 2,
                "templateId" : "6093869255ddb8004538cd66"
            }
            // console.log(executor_arr[j]);
        }
        // let query_executor = {templateId: ObjectID("6093869255ddb8004538cd66"),$and : [{name: {'$regex': executor}}]};
        // let executor_id = await run(query_executor);
        // console.log(executor_id);
        // console.log(executor_id);


        let openpl = sh.getRow(i).getCell(6).value;
        let closepl = sh.getRow(i).getCell(7).value;
        let openfact = sh.getRow(i).getCell(20).value;
        let closefact = sh.getRow(i).getCell(21).value;
        if(openpl){
            openpl=openpl.substr(3,2)+"."+openpl.substr(0,2)+openpl.substr(5,10);
        }
        if(closepl){
            closepl=closepl.substr(3,2)+"."+closepl.substr(0,2)+closepl.substr(5,10);
        }
        if(openfact){
            openfact=openfact.substr(3,2)+"."+openfact.substr(0,2)+openfact.substr(5,10);
        }
        if(closefact){
            closefact=closefact.substr(3,2)+"."+closefact.substr(0,2)+closefact.substr(5,10);
        }

        // console.log'startgjdfh'+(new Date(openpl));
        let status = sh.getRow(i).getCell(19).value+"";
        let query_status = {templateId: ObjectID("6093879d55ddb8004538cd67"),$and : [{name: {'$regex': status}}]};
        let status_id = await run(query_status);
        // console.log(status_id);

        let content = sh.getRow(i).getCell(4).value+"";
        let comment = sh.getRow(i).getCell(18).value+"";


        let create_date = sh.getRow(i).getCell(3).value;
        // console.log();
        if(create_date){
            create_date=create_date.substr(3,2)+"."+create_date.substr(0,2)+create_date.substr(5,10);
        }

        let query_mounth = {templateId: ObjectID("60ab330b55ddb8004538cd68")};
        let create_month_id = await run(query_mounth);
        // console.log(create_month_id[parseInt(create_date.substr(0,2))-1]);

        // let reconcile_Surname = 
        arr[i]={
            "templateId":ObjectID("609246b055ddb8004538cd62"),
            "name":application_number,
            "author" : {
                "_id" : author_id[0]._id,
                "passportType" : 2,
                "templateId" : "6093860155ddb8004538cd64"
            },
            // "templateId":ObjectId("609246b055ddb8004538cd62"),
            "reconcile" : reconcile_arr,
            "decision" : {
                "_id" : "6093871f26dde40009c54e7d",
                "passportType" : 2,
                "templateId" : "6093867855ddb8004538cd65"
            },
            "performer" : {
                "_id" : author_id[0]._id,
                "passportType" : 2,
                "templateId" : "6093860155ddb8004538cd64"
            },
            "area" : area,
            "number_disconnect" : number_disconnect,
            "mgd" : mgd,
            "chgd" : chgd,
            "detsad" : detsad,
            "hospital" : hospital,
            "school" : school,
            "other" : other,
            "executor" : executor_arr,
            "executor_text" : executor,
            "openpl" : openpl ? 'startgjdfh'+(new Date(openpl)).toISOString()+'gjdfh' : null,
            "closepl" : closepl ? 'startgjdfh'+(new Date(closepl)).toISOString()+'gjdfh': null,
            "closefact" : closefact ? 'startgjdfh'+(new Date(closefact)).toISOString()+'gjdfh' : null,
            "status" : {
                "_id" : status_id[0]._id,
                "passportType" : 2,
                "templateId" : "6093879d55ddb8004538cd67"
            },
            "amap" : {
                "lat" : null,
                "lng" : null,
                "id" : "",
                "comment" : ""
            },
            "creatdate" : create_date ? 'startgjdfh'+(new Date(create_date)).toISOString()+'gjdfh' : null,
            "create_year" : create_date.substr(6,4),
            "create_month" : {
                "_id" : create_month_id[parseInt(create_date.substr(0,2))-1]._id,
                "passportType" : 2,
                "templateId" : "60ab330b55ddb8004538cd68"
            },
            "create_day" : parseInt(create_date.substr(3,2)),
            "metadata" : {
                "docs" : {
                    "docs" : []
                },
                "photo" : {
                    "photo" : []
                },
                "video" : {
                    "video" : []
                }
            },
            "createDate" : 'startgjdfh'+(new Date).toISOString()+'gjdfh',
            "updateDate" : 'startgjdfh'+(new Date).toISOString()+'gjdfh',
            "isDeleted" : false,
            "passportType" : 1,
            "createInfo" : {
                "createDate" : 'startgjdfh'+(new Date).toISOString()+'gjdfh'
            },
            "comment" : comment,
            "content" : content,
            "Address" : "",
        };
        // console.log(arr[i]);
        // console.log('1', sh.getRow(i).getCell(1).value);//1

        // console.log(sh.getRow(i).getCell(2).value);
        // executor_arr=JSON.stringify(executor_arr);
        application_number=parseInt(application_number);
        // console.log(executor_arr);
        // console.log(application_number);
        update(application_number,openfact);
    }
    // var json = JSON.stringify(arr);
    // var fs = require('fs');
    // // // fs.writeFile('myjsonfile.json', json, 'utf8', callback);
    // fs.writeFile ("input3.json", json, function(err) {
    //     if (err) throw err;
    //     console.log('complete');
    //     }
    // );
    // console.log(arr);
    // for(i=2;i<=605;i++){
        // console.log(arr[i]);
        // addData(arr);
    // }
    

});
async function update(application_number,openfact) {
    try {
        await mongoClient.connect();
        const db = mongoClient.db("agisads");
        const collection = db.collection("facilityPassports");
        console.log(application_number);
        // application_number=Number(application_number);
        // const result2 = await collection.find({templateId: ObjectID("609246b055ddb8004538cd62"),name:application_number}).toArray();
        // console.log(result2);
        // const name = await  collection.find(query).toArray();
        // const name = await  collection.updateMany({templateId: ObjectID("6093860155ddb8004538cd64"),$and : [{name: {'$regex': str}}, {name: {'$regex': str2}}]}).toArray();
        const result = await collection.updateOne({templateId: ObjectID("609246b055ddb8004538cd62"),name:application_number}, { $set: {openfact: new Date(openfact)}});

        // const result = await collection.updateOne({templateId: ObjectID("609246b055ddb8004538cd62"),name:application_number}, { $set: {executor: executor_arr}});
        // const name = await  collection.find({templateId: ObjectID("6093860155ddb8004538cd64"),name: {'$regex': str}}).toArray();
        console.log(result);
        // console.log(name);
        // return name;
         
    }catch(err) {
        console.log(err);
    } 
}



