// Requiring express to handle routing
const express = require("express");

// The fileUpload npm package for handling
// file upload functionality
const fileUpload = require("express-fileupload");
const { createRunSheet } = require("./processFile");

const PORT = process.env.PORT || 3001;

// Creating app
const app = express();

// Passing fileUpload as a middleware
app.use(fileUpload());
var uploadedFile = null
var uploadFileName = null


// For handling the upload request
app.post("/upload", function (req, res) {

    // When a file has been uploaded
    if (req.files && Object.keys(req.files).length !== 0 && req.files.uploadFile.mimetype == 'text/csv') {
        
        // Uploaded path
        uploadedFile = req.files.uploadFile;

        // Logging uploading file
        console.log(uploadedFile);

        // Upload path
        const uploadPath = "./input/" + uploadedFile.name;

        // To save the file using mv() function
        uploadedFile.mv(uploadPath, function (err) {
        if (err) {
            console.log(err);
            res.send("Failed !!");
        } else {
            uploadFileName=uploadedFile.name;
            createRunSheet(uploadPath,uploadedFile.name);
            res.send("Successfully Uploaded !!")
        };
    });
    } else res.send("Failed to upload file! Only CSV files are accepted.");
});

// To handle the download file request
app.get("/download", function (req, res) {

    // The res.download() talking file path to be downloaded
    res.download('./output/'+uploadFileName.replace('.csv','')+'-RunSheet.xlsx', function (err) {
        if (err) {
        console.log(err);
        };
        uploadFileName = null;

    });
});

// GET request to the root of the app
app.get("/", function (req, res) {

// Sending index.html file as response to the client
res.sendFile(__dirname + "/index.html");
});

// Makes app listen to port 3000
// app.listen(3000, function (req, res) {
// console.log("Started listening to port 3000");
// });
app.listen(PORT, function(req,res) {
    console.log(`Server listening on ${PORT}`);
});