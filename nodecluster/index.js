// jshint esversion: 6
var cluster = require("cluster");
var os = require("os");
var http = require("http");
const PORT = 9090;

var noOfCPUs = os.cpus().length;
var workers = [];


cluster.on('online', workerInit);
cluster.on('exit', workerExit);

if(cluster.isMaster) {
    //main process
    for (let index = 0; index < noOfCPUs; index++) {
        var worker = cluster.fork();
        if(worker)
            workers.push(worker);
    }
}
else {
    //worker items
    http.createServer(function(req, res) {
        res.writeHead(200);
        res.end('process ' + process.pid + 'says hello');
    }).listen(PORT);
}

//Called when worker initiated
function workerInit(worker) {
    console.log("Worker " + worker.process.pid + " initialized " + " (listening at " + PORT + ") and live now");
}

//Called when worker exited
function workerExit(worker, code, signal) {
    console.log("Worker " + worker.process.pid + " exited " + " (was listening at " + PORT + ") and live now");
    for (let index = workers-1; index >= 0; index--) {
        var workerProcess = workers[index];
        if(workerProcess.process.pid === worker.process.pid) {
            //now delete the worker from the stack as worker is not alive any more
            workers.splice(index, 1);
        }
    }
}
