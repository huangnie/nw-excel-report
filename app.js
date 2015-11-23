 var format = function(time, format){
    var t = new Date(time);
    var tf = function(i){return (i < 10 ? '0' : '') + i};
    return format.replace(/yyyy|MM|dd|HH|mm|ss/g, function(a){
	    switch(a){
	        case 'yyyy':
	        return tf(t.getFullYear());
	        break;
	        case 'MM':
	        return tf(t.getMonth() + 1);
	        break;
	        case 'mm':
	        return tf(t.getMinutes());
	        break;
	        case 'dd':
	        return tf(t.getDate());
	        break;
	        case 'HH':
	        return tf(t.getHours());
	        break;
	        case 'ss':
	        return tf(t.getSeconds());
	        break;
	    }
    })
} 

var report = new Vue({
		  el: '#report',
		  data: {
		    levels: ['绝','机','密','平'],
		    to_user:"",
		    from_date: format(new Date().getTime(), 'yyyy-MM-dd'),
		    mailer_no:"",
		    inform_no:"",
		    from_user:"",
			from_tel:"",
			tip:"",  
		    report:[
		    	{ 
		    		file:  { name: '', level: '绝', num:1 },
		      		margize: { name: '',　level: '机',　num:1　}
		      	},
		    	{ 
		    		file:  { name: '', level: '密', num:1 },
		      		margize: { name: '', level: '平', num:1 }
		      	},
		      	{ 
		    		file:  { name: '', level: '绝', num:1 },
		      		margize: { name: '',　level: '机',　num:1　}
		      	},
		    	{ 
		    		file:  { name: '', level: '密', num:1 },
		      		margize: { name: '', level: '平', num:1 }
		      	},
		      	{ 
		    		file:  { name: '', level: '绝', num:1 },
		      		margize: { name: '',　level: '机',　num:1　}
		      	},
		    	{ 
		    		file:  { name: '', level: '密', num:1 },
		      		margize: { name: '', level: '平', num:1 }
		      	}
		    ]  
		  },
		  methods: {
			    add: function (e) {   
					this.report.push({ 
			    		file:  { name: '', level: '绝', num:1 },
			      		margize: { name: '', level: '机', num:1 }
			      	});
				},

				del: function () {
					var len = this.report.length;
					console.log(len);
					this.report = this.report.splice(0,len-1);
				},

				save: function (){
					var fs = require("fs");
					var path = require("path");
					var Excel = require("exceljs");
					var workbook = new Excel.Workbook();
					//Set Workbook Properties
					workbook.creator = "huangnie";
					workbook.lastModifiedBy = "wangcheng";
					workbook.created = new Date();
					workbook.modified = new Date();

					var worksheet = workbook.addWorksheet("My Sheet");	

					worksheet.getColumn("A").width = 50;
					worksheet.getColumn("B").width = 5;
					worksheet.getColumn("C").width = 5;

					worksheet.getColumn("D").width = 50;
					worksheet.getColumn("E").width = 5;
					worksheet.getColumn("F").width = 5;

					worksheet.mergeCells("A1:C1");
					worksheet.getCell("A1").value = " 单位：" + this.to_user;
					worksheet.mergeCells("D1:F1");
					worksheet.getCell("D1").value = " 发报时间："+ this.from_date+"　　　信封号："+this.mailer_no+"　　　通知单号："+this.inform_no;
					worksheet.getRow(1).commit(); // now rows 1 and two are committed.
					
					worksheet.getCell("A2").value = "文件";
					worksheet.getCell("A2").alignment = { vertical: "middle", horizontal: "center" };
					worksheet.getCell("B2").value = "密级";
					worksheet.getCell("B2").alignment = { vertical: "middle", horizontal: "center" };
					worksheet.getCell("C2").value = "份数";
					worksheet.getCell("C2").alignment = { vertical: "middle", horizontal: "center" };

					worksheet.getCell("D2").value = "内刊";
					worksheet.getCell("D2").alignment = { vertical: "middle", horizontal: "center" };
					worksheet.getCell("E2").value = "密级";
					worksheet.getCell("E2").alignment = { vertical: "middle", horizontal: "center" };
					worksheet.getCell("F2").value = "份数";
					worksheet.getCell("F2").alignment = { vertical: "middle", horizontal: "center" };

					var save_report = this.report;
					var num=2;	

					for(var i in save_report) {
						num ++;
						worksheet.getCell("A"+num).value = " " + save_report[i].file.name;
						worksheet.getCell("B"+num).value = save_report[i].file.level;
						worksheet.getCell("B"+num).alignment = { vertical: "middle", horizontal: "center" };
						worksheet.getCell("C"+num).value = save_report[i].file.num;
						worksheet.getCell("C"+num).alignment = { vertical: "middle", horizontal: "center" };
						 
						worksheet.getCell("D"+num).value = " " + save_report[i].margize.name;
						worksheet.getCell("E"+num).value = save_report[i].margize.level;　
						worksheet.getCell("E"+num).alignment = { vertical: "middle", horizontal: "center" };
						worksheet.getCell("F"+num).value = save_report[i].margize.num;
						worksheet.getCell("F"+num).alignment = { vertical: "middle", horizontal: "center" };
					}	    

					num ++;
					worksheet.mergeCells("A"+num+":F"+num);
					worksheet.getCell("A"+num).value = this.tip;
					worksheet.getCell("A"+num).alignment = { vertical: "middle", horizontal: "center" };
					
					num ++;
					worksheet.mergeCells("A"+num+":C"+num);
					worksheet.getCell("A"+num).value = this.from_user + " 制表";
					worksheet.mergeCells("D"+num+":F"+num);
					worksheet.getCell("D"+num).value = "电话："+ this.from_tel +" ";
					worksheet.getCell("D"+num).alignment = { vertical: "middle", horizontal: "right" };
					worksheet.getRow(1).commit(); // now rows 1 and two are committed.

					function mkdir(data_dir){
						if (!fs.existsSync(data_dir)) {
	  				        parent_dir = path.dirname(data_dir);
	  				        if (!fs.existsSync(data_dir)) {
								mkdir(parent_dir);
							}

							fs.mkdirSync(data_dir);
	  				       	return true;
	  				    }
	  				    return true;
					} 

					// data_dir = path.dirname(process.cwd())+'/data/excel/';
					data_dir = 'D:/report/data/excel/';
					if (!fs.existsSync(data_dir)) {
  				        mkdir(data_dir);
  				        console.log('Common目录创建成功');
  				    }

  				    filename = this.to_user　+ "_" +　this.from_date;
					filepath　= data_dir + filename +　'.xlsx';
					isSave = true;
					if (fs.existsSync(filepath)) {
  				        isSave = confirm('将会覆盖已存在的文档，确认要覆盖吗');
  				    }
					if(isSave){
						workbook.xlsx.writeFile(filepath).then(function() { 
							alert("已保存位置: " + filepath);
						});

						// filepath = data_dir + filename + '.cvs';
						// workbook.csv.writeFile(filepath).then(function() { 
						// 	// alert("已保存位置: " + filepath);
						// });
					}
				},
			},
		});

	
		$("input").mouseover(function(){
		  	$(this).focus(); 
		  	if($(this).createTextRange){  
		       var r= $(this).createTextRange();  
		       r.moveStart('character', field.value.length);  
		       // r.collapse();  
		       // r.select();//加上这一句是使文本框里的文字为选中状态  
		    }  
		});

		$("input").mouseout(function(){
		  	$(this).blur(); 
		});

		$("#add_or_del").mouseover(function(){
		  	if($("#group").attr('class') == 'hide'){
		  		$("#group").attr('class','show');
		  		$("#group").show(300); 
		  	}
		});

		$("#add_or_del").mouseout(function(){
		  	$("#group").attr('class','hide');
		  	$("#group").hide(500); 
		});


        var gui = require('nw.gui');
        var win = gui.Window.get();
      //  win.enterFullscreen();
          win.leaveFullscreen();

	    var win = gui.Window.get()
	    win.on('resize', function(){
	      // alert(1)   //没有触发
	      // $('body').width(win.width).height(win.height)
	    })