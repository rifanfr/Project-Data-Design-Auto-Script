#target illustrator

/*
BOX DIELINE GENERATOR - ONEDRIVE JSON
Output identik dengan V1 final. Perbedaannya hanya sumber data:
Length, Width, Height, Flute, Posisi Kuping dibaca dari design_data_*.json terbaru.
Glue flap is present only along the middle panel row.
*/

(function () {
    function mm(v) { return v * 2.834645669; }
    function fmt(v) { return String(Math.round(v*100)/100); }
    function parseNum(s) { return parseFloat(String(s).replace(/\s/g,"").replace(",",".")); }
    function cleanText(v) { return (v === undefined || v === null) ? "" : String(v).replace(/^\s+|\s+$/g, ""); }

    // Color follows the document color space (CMYK or RGB)
    function makeColor(docSpace,c,m,y,k,r,g,b) {
        if (docSpace==DocumentColorSpace.CMYK) {
            var cc=new CMYKColor();
            cc.cyan=c; cc.magenta=m; cc.yellow=y; cc.black=k;
            return cc;
        }
        var rc=new RGBColor();
        rc.red=r; rc.green=g; rc.blue=b;
        return rc;
    }
    var BLACK, RED;

    function calculateFormula(length,width,height,flute) {
        var add=0, kupingan=32, middleAdd=0;
        if (flute=="B") { add=3; kupingan=32; middleAdd=5; }
        else if (flute=="C") { add=4; kupingan=32; middleAdd=7; }
        else if (flute=="BC") { add=7; kupingan=35; middleAdd=15; }
        else { throw new Error("Flute harus B, C, atau BC."); }

        var P1=length+add;
        var L1=width+add;
        var P2=P1;
        var L2=L1-3;
        var middle=height+middleAdd;
        var isEven=(Math.round(L1)%2===0);
        var flap;
        if (flute=="B") flap=(L1+(isEven?0:1))/2;
        else if (flute=="C") flap=(L1+(isEven?2:1))/2;
        else flap=(L1+(isEven?4:3))/2;

        return {
            kupingan:kupingan, P1:P1, L1:L1, P2:P2, L2:L2,
            flapTop:flap, middle:middle, flapBottom:flap,
            totalWidth:kupingan+P1+L1+P2+L2,
            totalHeight:flap+middle+flap
        };
    }

    // ============================================================
    // INPUT FROM ONEDRIVE JSON
    // ============================================================

    var targetFolderName="DESIGN DATA PWR AT Sharepoint";
    var dataFolder=null;
    var rootsToTry=[];

    function pushRoot(pathText) {
        if (!pathText) return;
        var f=new Folder(pathText);
        if (f.exists) rootsToTry.push(f);
    }
    try { pushRoot($.getenv("OneDriveCommercial")); } catch (e0) {}
    try { pushRoot($.getenv("OneDrive")); } catch (e1) {}
    try { pushRoot($.getenv("UserProfile")); } catch (e2) {}
    pushRoot(new Folder("~").fsName);

    for (var rr=0; rr<rootsToTry.length && !dataFolder; rr++) {
        var direct=new Folder(rootsToTry[rr].fsName+"/"+targetFolderName);
        if (direct.exists) { dataFolder=direct; break; }
        var children=[];
        try { children=rootsToTry[rr].getFiles(function (f) { return f instanceof Folder; }); } catch (e3) {}
        for (var cc=0; cc<children.length; cc++) {
            var nested=new Folder(children[cc].fsName+"/"+targetFolderName);
            if (nested.exists) { dataFolder=nested; break; }
        }
    }

    if (!dataFolder) {
        alert("Folder OneDrive belum terdeteksi otomatis.\n\nPilih folder:\n"+targetFolderName);
        var picked=Folder.selectDialog("Pilih folder "+targetFolderName);
        if (!picked) return;
        if (picked.name!==targetFolderName) {
            var child=new Folder(picked.fsName+"/"+targetFolderName);
            if (child.exists) picked=child;
        }
        dataFolder=picked;
    }

    var jsonFiles=dataFolder.getFiles(function (f) {
        return (f instanceof File) && /^design_data_.*\.json$/i.test(f.name);
    });
    if (!jsonFiles || jsonFiles.length===0) {
        alert("Tidak ada design_data_*.json di:\n"+dataFolder.fsName+"\n\nPastikan OneDrive sudah selesai sync.");
        return;
    }

    var dataFile=jsonFiles[0];
    for (var jf=1; jf<jsonFiles.length; jf++) {
        if (jsonFiles[jf].modified>dataFile.modified) dataFile=jsonFiles[jf];
    }

    dataFile.encoding="UTF-8";
    if (!dataFile.open("r")) {
        alert("JSON terbaru tidak bisa dibuka:\n"+dataFile.fsName);
        return;
    }
    var raw=dataFile.read();
    dataFile.close();
    raw=raw.replace(/^\uFEFF/,"");

    var data;
    try { data=(typeof JSON!=="undefined" && JSON.parse) ? JSON.parse(raw) : eval("("+raw+")"); }
    catch (jsonErr) {
        alert("Format JSON tidak valid:\n"+jsonErr+"\n\nFile: "+dataFile.fsName);
        return;
    }
    if (!data || typeof data!=="object") {
        alert("Isi JSON tidak berisi data yang valid:\n"+dataFile.name);
        return;
    }

    var length=parseNum(data.length);
    var width=parseNum(data.width);
    var height=parseNum(data.height);
    if (isNaN(length)||isNaN(width)||isNaN(height)||length<=0||width<=0||height<=0) {
        alert("Data dimensi JSON tidak valid.\n\nLength: "+data.length+"\nWidth: "+data.width+"\nHeight: "+data.height);
        return;
    }

    var flute=cleanText(data.flute).toUpperCase();
    flute=flute.replace(/\s/g,"").replace(/\/F$/,"");
    if (!(flute=="B" || flute=="C" || flute=="BC")) {
        alert("Flute JSON tidak didukung: "+cleanText(data.flute)+"\n\nGunakan B, C, atau BC.");
        return;
    }

    var sideRaw=cleanText(data.posisi_kuping).toLowerCase();
    var side=(sideRaw=="kiri") ? "Kiri" : ((sideRaw=="kanan") ? "Kanan" : "");
    if (!side) {
        alert("posisi_kuping JSON harus Kiri atau Kanan.\nNilai saat ini: "+cleanText(data.posisi_kuping));
        return;
    }

    // ============================================================
    // GENERATE (sama dengan V1)
    // ============================================================

    try {
    var result=calculateFormula(length,width,height,flute);

    var widths=[result.kupingan,result.P1,result.L1,result.P2,result.L2];
    if (side=="Kanan") {
        widths=[result.L2,result.P2,result.L1,result.P1,result.kupingan];
    }
    var heights=[result.flapTop,result.middle,result.flapBottom];

    var doc=app.documents.length ? app.activeDocument : app.documents.add(DocumentColorSpace.CMYK,mm(2000),mm(1000));

    BLACK=makeColor(doc.documentColorSpace,0,0,0,100,0,0,0);
    RED=makeColor(doc.documentColorSpace,0,100,100,20,180,30,30);

    // Case-insensitive lookup so "Dimensions" and "DIMENSIONS" are the same layer
    function getLayer(name) {
        for (var i=0;i<doc.layers.length;i++) {
            if (doc.layers[i].name.toLowerCase()==name.toLowerCase()) return doc.layers[i];
        }
        var l=doc.layers.add();
        l.name=name;
        return l;
    }
    var dielineLayer=getLayer("UKURAN");
    dielineLayer.locked=false;
    dielineLayer.visible=true;
    var dimensionLayer=getLayer("DIMENSIONS");
    dimensionLayer.locked=false;
    dimensionLayer.visible=true;

    var totalWpt=mm(result.totalWidth);
    var totalHpt=mm(result.totalHeight);
    var ab=doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var cx=(ab[0]+ab[2])/2;
    var cy=(ab[1]+ab[3])/2;
    var startX=cx-totalWpt/2;
    var startY=cy+totalHpt/2;

    var xPositions=[];
    var x=startX;
    for (var col=0;col<widths.length;col++) {
        xPositions.push(x);
        var y=startY;
        for (var row=0;row<heights.length;row++) {
            var isGlueFlap=(side=="Kanan") ? col==widths.length-1 : col==0;
            if (!isGlueFlap || row==1) {
                var rect=dielineLayer.pathItems.rectangle(y,x,mm(widths[col]),mm(heights[row]));
                rect.stroked=true;
                rect.strokeWidth=1;
                rect.strokeColor=BLACK;
                rect.filled=false;
            }
            y-=mm(heights[row]);
        }
        x+=mm(widths[col]);
    }
    xPositions.push(x);

    function drawLine(x1,y1,x2,y2,w,color,layer) {
        var q=layer.pathItems.add();
        q.setEntirePath([[x1,y1],[x2,y2]]);
        q.stroked=true; q.strokeWidth=w; q.strokeColor=color; q.filled=false;
        return q;
    }

    function arrow(x,y,dx,dy,size) {
        var len=Math.sqrt(dx*dx+dy*dy);
        if (len===0) return;
        var ux=dx/len, uy=dy/len, px=-uy, py=ux;
        var bx=x-ux*size, by=y-uy*size;
        var q=dimensionLayer.pathItems.add();
        q.setEntirePath([[x,y],[bx+px*size*0.42,by+py*size*0.42],[bx-px*size*0.42,by-py*size*0.42],[x,y]]);
        q.closed=true; q.stroked=false; q.filled=true; q.fillColor=BLACK;
    }

    function addText(txt,x,y,rotation) {
        var t=dimensionLayer.textFrames.add();
        t.contents=txt;
        t.textRange.characterAttributes.size=40;
        t.textRange.characterAttributes.fillColor=RED;
        try { t.textRange.characterAttributes.textFont=app.textFonts.getByName("SourceSans3-Regular"); } catch(e) {}
        t.left=x-t.width/2;
        t.top=y+t.height/2;
        if (rotation) t.rotate(rotation,true,true,true,true,Transformation.CENTER);
        return t;
    }

    function hDim(x1,x2,y,label,textGap) {
        drawLine(x1,y,x2,y,1,BLACK,dimensionLayer);
        arrow(x1,y,1,0,mm(3));
        arrow(x2,y,-1,0,mm(3));
        addText(label,(x1+x2)/2,y+mm(textGap),0);
    }

    function vDim(x,y1,y2,label,textGap,textSide) {
        drawLine(x,y1,x,y2,1,BLACK,dimensionLayer);
        arrow(x,y1,0,1,mm(3));
        arrow(x,y2,0,-1,mm(3));
        var textX=(textSide=="LEFT") ? x-mm(textGap) : x+mm(textGap);
        addText(label,textX,(y1+y2)/2,(textSide=="LEFT")?90:-90);
    }

    var top=startY, bottom=startY-totalHpt, left=startX, right=startX+totalWpt;
    var GAP=8, SEG_LEN=20, TOTAL_LEN=40;
    var yStart=top+mm(GAP);
    var yOverall=yStart+mm(TOTAL_LEN);
    drawLine(left,yStart,left,yOverall,0.75,BLACK,dimensionLayer);
    drawLine(right,yStart,right,yOverall,0.75,BLACK,dimensionLayer);
    hDim(left,right,yOverall,fmt(result.totalWidth)+" mm",7);

    var ySegments=yStart+mm(SEG_LEN);
    for (var i=0;i<xPositions.length;i++) drawLine(xPositions[i],yStart,xPositions[i],ySegments,0.75,BLACK,dimensionLayer);
    for (var j=0;j<widths.length;j++) hDim(xPositions[j],xPositions[j+1],ySegments,fmt(widths[j])+" mm",7);

    var y1=top;
    var y2=top-mm(result.flapTop);
    var y3=y2-mm(result.middle);
    var y4=bottom;
    var yList=[y1,y2,y3,y4];
    var dimensionOnLeft=(side=="Kanan");
    var dir=dimensionOnLeft ? -1 : 1;
    var textSide=dimensionOnLeft ? "LEFT" : "RIGHT";
    var baseX=(dimensionOnLeft ? left : right)+dir*mm(GAP);
    var xSeg=baseX+dir*mm(SEG_LEN);
    var xOverall=baseX+dir*mm(TOTAL_LEN);

    for (var k=0;k<yList.length;k++) drawLine(baseX,yList[k],xSeg,yList[k],0.75,BLACK,dimensionLayer);
    vDim(xSeg,y2,y1,fmt(result.flapTop)+" mm",8,textSide);
    vDim(xSeg,y3,y2,fmt(result.middle)+" mm",8,textSide);
    vDim(xSeg,y4,y3,fmt(result.flapBottom)+" mm",8,textSide);
    drawLine(baseX,top,xOverall,top,0.75,BLACK,dimensionLayer);
    drawLine(baseX,bottom,xOverall,bottom,0.75,BLACK,dimensionLayer);
    vDim(xOverall,bottom,top,fmt(result.totalHeight)+" mm",8,textSide);

    alert("GENERATED FROM ONEDRIVE JSON\n\nJSON: "+dataFile.name+"\nNo. MG: "+cleanText(data.no_mg)+"\nSize: "+fmt(length)+" x "+fmt(width)+" x "+fmt(height)+" mm\n\nFlute: "+flute+"\nKupingan: "+side+" / "+fmt(result.kupingan)+" mm\nP1: "+fmt(result.P1)+" mm\nL1: "+fmt(result.L1)+" mm\nP2: "+fmt(result.P2)+" mm\nL2: "+fmt(result.L2)+" mm\n\nTotal: "+fmt(result.totalWidth)+" mm x "+fmt(result.totalHeight)+" mm");
        return true;
        } catch(err) {
            alert("Error: "+err.message+(err.line ? "\nBaris: "+err.line : ""));
            return false;
        }
    })();
