const http=require('http'),fs=require('fs'),path=require('path');
const root=__dirname;
http.createServer((req,res)=>{
  let p=req.url.split('?')[0]; if(p==='/')p='/index.html';
  const f=path.join(root,path.normalize(p).replace(/^(\.\.[\/\\])+/,''));
  fs.readFile(f,(e,d)=>{
    if(e){res.writeHead(404);return res.end('not found');}
    const ext=path.extname(f);
    const ct={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8',
      '.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8'}[ext]||'application/octet-stream';
    res.writeHead(200,{'Content-Type':ct,'Cache-Control':'no-store'});res.end(d);
  });
}).listen(3200,()=>console.log('alfa on 3200'));
