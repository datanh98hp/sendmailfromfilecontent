const express =  require('express')
const nodemailer = require("nodemailer");
const app = express();
const cors = require('cors');
app.use(cors({
    origin: '*'
}));
const readExcel = require('read-excel-file/node')
const multer = require("multer");

require('dotenv').config()

// Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
       cb(null, __dirname + '/uploads/')
    },
    filename: (req, file, cb) => {
       cb(null,file.originalname)
    }
});
const upload = multer({ storage:storage});
app.use(express.json())
app.get('/', (req, res)=>{
    console.log(__dirname)
    res.send('OK')
});
/// readfile xlsx
app.post("/upload_files", upload.single("file"), uploadFiles);
async function uploadFiles(req, res) {
    console.log(req.body);
    console.log(req.files);
    if (!req.file) {
        return res.json({message:"miss file upload",error:400})
    }
    const data = await readDataFromFileXLSX(__dirname+'/uploads/'+req.file.filename);

    /// send mail
    data.map((item)=>{
        if (item.email!==null) {
            sendMail({
                from:'dat198hp@gmail.com',
                to:item.email,
                subject:`LUONG`,
                text:`Thong tin luong`,
                html:`
                    <div>${item.so_gio_lam }</div>
                    <div>${item.so_ngay_nghi}</div>
                    <div>${item.so_pc}</div>
                    <div>${item.tong_luong}</div>
                `
            });
        }
    });

    res.json({ message: "Successfully uploaded files" });
}


async function readDataFromFileXLSX(pathFile){

    const data = await readExcel(pathFile);
    const listEmail = [];
    const dataMail = [];

    for(i in data){
        if(i>1){ // loai đi ten cot
           let name = data[i][1];
           let email = data[i][3];
           let so_gio_lam = data[i][4];
           let so_gio_ot = data[i][5];
           let so_pc = data[i][6];
           let so_ngay_nghi = data[i][7];
           let tong_luong = data[i][8];

           dataMail.push({
            name,email,so_gio_lam,so_gio_ot,so_pc,so_ngay_nghi,tong_luong
           })
           
        }
       
    }
    return dataMail;
    //return res.json(dataMail);

}
// send mail
async function sendMail(content){
    const {from,to,subject,text,html} = content;

    const transporter = nodemailer.createTransport({
        service:"gmail",
     //    service: 'gmail',
        host: 'smtp.gmail.com',
        port: 465,
        secure: true,
         auth: {
           // TODO: replace `user` and `pass` values from <https://forwardemail.net>
           user:process.env.FROM_EMAIL , //'dat198hp@gmail.com', // env config
           pass:process.env.PASS_EMAIL_APP//'rlbjimkavjwiihee'     // env config
         }
       });

       // send mail with defined transport object
        await transporter.sendMail({
           from: content.from, //'dat198hp@gmail.com', // sender address
           to: `${content.to}`, // list of receivers
           subject: content.subject, // Subject line
           text: content.text, // plain text body
           html: content.html, // html body
         },(err)=>{
             if (err) {
                 return res.json(err);
             }
             return res.json('success');
         });

        return true;
         
}

app.post('/read-file', async (req,res)=>{

    const data = await readExcel('./test-file.xlsx');
    const listEmail = [];
    const dataMail = [];

    console.log(req.files)
    // for(i in data){
    //     if(i>1){ // loai đi ten cot
    //        let name = data[i][1];
    //        let email = data[i][3];
    //        let so_gio_lam = data[i][4];
    //        let so_gio_ot = data[i][5];
    //        let so_pc = data[i][6];
    //        let so_ngay_nghi = data[i][7];
    //        let tong_luong = data[i][8];

    //        dataMail.push({
    //         name,email,so_gio_lam,so_gio_ot,so_pc,so_ngay_nghi,tong_luong
    //        })
           
    //     }
       
    // }
    //return res.json(dataMail);
}); 


//
app.post("/send-mail", async (req,res)=>{

    const {email} = req.body

    const transporter = nodemailer.createTransport({
       service:"gmail",
    //    service: 'gmail',
       host: 'smtp.gmail.com',
       port: 465,
       secure: true,
        auth: {
          // TODO: replace `user` and `pass` values from <https://forwardemail.net>
          user: 'dat198hp@gmail.com',
          pass: 'rlbjimkavjwiihee'
        }
      });

  
      // send mail with defined transport object
       await transporter.sendMail({
          from: 'dat198hp@gmail.com', // sender address
          to: `${email}`, // list of receivers
          subject: "Test sendmail✔", // Subject line
          text: "Hello world? from Da", // plain text body
          html: "<b>Hello world?</b>", // html body

        },(err)=>{
            if (err) {
                return res.json(err);
            }
            return res.json('success');
        });
      
       
        // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>
      
        //
        // NOTE: You can go to https://forwardemail.net/my-account/emails to see your email delivery status and preview
        //       Or you can use the "preview-email" npm package to preview emails locally in browsers and iOS Simulator
        //       <https://github.com/forwardemail/preview-email>


})

const port = process.env.PORT || 8080;

app.listen(port,()=>{
    console.log(`server running on port : ${port}`)
})

