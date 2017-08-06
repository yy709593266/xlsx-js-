// 算平均值时总时间的被除数以及工作日的被除数
var allDaysNum = 1;
var workDaysNum = 1;

// 存储当前所有的日期以及所有工作日的日期
var saveAllDays = [];
var saveWorkDays = [];

// 表格名称
var labTitle = '';

// 建立对象存储显示数据，包括姓名，工作日总时长，工作日平均时长，总时长，总平均时长
var allData = [];

// 所有人的总时间平均值
var allmean = 0;
// 所有人工作日的时间平均值
var workmean = 0;
// x轴显示不同人的姓名
var allPerpleName = [];
// 所有人的总时长
var allPeopleAllTime = [];
// 所有人的工作日时长
var allPerpleWorkTime = [];
// 所有人的总平均时长
var allPerpleAllAvgTime = [];
// 所有人的工作日平均时长
var allPerpleWorkAvgTime = [];


$("#excel-file").change(function(e) {
    var files = e.target.files;
    // 获取图标表示的具体实验室名称
    var pathArr = this.value.split("\\");
    labTitle = pathArr[pathArr.length - 1];
    pathArr = labTitle.split(".");
    labTitle = pathArr[0];

    var fileReader = new FileReader();
    fileReader.onload = function(ev) {
        try {
            var data = ev.target.result,
                workbook = XLSX.read(data, {
                    type: 'binary'
                }), // 以二进制流方式读取得到整份excel表格对象
                persons = []; // 存储获取到的数据
        } catch (e) {
            console.log('文件类型不正确');
            return;
        }

        // 表格的表格范围，可用于判断表头是否数量是否正确
        var fromTo = '';
        // 遍历每张表读取
        for (var sheet in workbook.Sheets) {
            if (workbook.Sheets.hasOwnProperty(sheet)) {
                fromTo = workbook.Sheets[sheet]['!ref'];
                persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
            	// 建立用来绘制图表的数据
            	getDrawData(persons);
            	
            }
        }
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);

    function getDrawData(persons){
    	allData = [];
    	saveAllDays = [];
    	saveWorkDays = [];
    	allDaysNum = 1;
    	workDaysNum = 1;
    	allmean = 0;
    	workmean = 0;
    	allPerpleName = [];
    	allPeopleAllTime = [];
    	allPerpleWorkTime = [];
    	allPerpleAllAvgTime = [];
    	allPerpleWorkAvgTime = [];
    	
        var personsLen = persons.length;
        var allDataName = [];

        for(var i = 0; i < personsLen; i++) {
        	// 将左右时间和工作日时间存储起来
        	saveAllDays.push(persons[i]["考勤日期"]);
        	if(!isWeekDay(persons[i]["考勤日期"])) {
        		saveWorkDays.push(persons[i]["考勤日期"]);
        	}
        	if(allData.length == 0) {
        		var firstData = {};
        			firstData.name = persons[0]["姓名"];
        			firstData.allTime = parseFloat(persons[0]["有效时长"].split(':').join('.'));
        			if(isWeekDay(persons[0]["考勤日期"])) {
        				firstData.workTime = 0;
        			}else {
        				firstData.workTime = parseFloat(persons[0]["有效时长"].split(':').join('.'));
        			}
        			allData.push(firstData);
        	}else {
        		for(var j = 0; j < allData.length; j++) {
        			if(allDataName.indexOf(allDataName[j]) == -1) {
        				allDataName.push(allData[j].name);
        			}
            	}
            	if(allDataName.indexOf(persons[i]["姓名"]) == -1) {
            		var newData = {};
            			newData.name = persons[i]["姓名"];
            			newData.allTime = parseFloat(persons[i]["有效时长"].split(':').join('.'));
            			if(isWeekDay(persons[i]["考勤日期"])) {
            				newData.workTime = 0;
            			}else {
            				newData.workTime = parseFloat(persons[i]["有效时长"].split(':').join('.'));
            			}
            			allData.push(newData);
            	}else {
            		var idx = allDataName.indexOf(persons[i]["姓名"]);
            		allData[idx].allTime += parseFloat(persons[i]["有效时长"].split(':').join('.'));
            		if(isWeekDay(persons[i]["考勤日期"])) {
            			allData[idx].workTime += 0;
            		}else {
            			allData[idx].workTime += parseFloat(persons[i]["有效时长"].split(':').join('.'));
            		}
            	}
        	}
        }

        // 获取最大日期的那一天和最小日期那一天，然后求得总的天数
        var allMaxDay = saveAllDays[0];
        var allMinDay = saveAllDays[0];
        for(var i = 0; i < saveAllDays.length; i++) {
        	if(parseInt(allMaxDay.split('/').join('')) < parseInt(saveAllDays[i].split('/').join(''))) {
				allMaxDay = saveAllDays[i];
			}
			if(parseInt(allMinDay.split('/').join('')) > parseInt(saveAllDays[i].split('/').join(''))) {
				allMinDay = saveAllDays[i];
			}
        }
        allDaysNum = (Date.parse(allMaxDay) - Date.parse(allMinDay)) / (1000 * 60 * 60 * 24) + 1;

        var workMaxDay = saveWorkDays[0];
        var workMinDay = saveWorkDays[0];
        for(var i = 0; i < saveWorkDays.length; i++) {
        	if(parseInt(workMaxDay.split('/').join('')) < parseInt(saveWorkDays[i].split('/').join(''))) {
				workMaxDay = saveWorkDays[i];
			}
			if(parseInt(workMinDay.split('/').join('')) > parseInt(saveWorkDays[i].split('/').join(''))) {
				workMinDay = saveWorkDays[i];
			}
        }
        workDaysNum = (Date.parse(workMaxDay) - Date.parse(workMinDay)) / (1000 * 60 * 60 * 24) + 1;
        
        // 计算每个人的平均时长
        for(var i = 0; i < allData.length; i++) {
        	allData[i].allAvgTime = (allData[i].allTime/allDaysNum);
        	allData[i].workAvgTime = (allData[i].workTime/workDaysNum);
        }
        // 将每个人的姓名，时间等提出来作为highcharts的参数
        for(var i = 0; i < allData.length; i++) {
        	allPerpleName.push(allData[i].name);
        	allPeopleAllTime.push(allData[i].allTime);
        	allPerpleWorkTime.push(allData[i].workTime);
        	allPerpleAllAvgTime.push(allData[i].allAvgTime);
        	allPerpleWorkAvgTime.push(allData[i].workAvgTime);
        }
        // 总体的工作日平均值
	    var workmeanplus = 0;
	    for(var i = 0; i < allPerpleWorkAvgTime.length; i++) {
	    	workmeanplus += allPerpleWorkAvgTime[i];
	    }
	    workmean = parseFloat(workmeanplus / (allPerpleName.length));

	    // 总体的平均值
	    var allmeanplus = 0;
	    for(var i = 0; i < allPerpleAllAvgTime.length; i++) {
	    	allmeanplus += allPerpleAllAvgTime[i];
	    }
	    allmean = parseFloat(allmeanplus / (allPerpleName.length));
    }

    var timer = setTimeout(function(){
    	var chartConfig = {
			chart: {
		        type: 'column'
		    },
		    xAxis: {
		        categories: allPerpleName
		    },
		    credits: {
		    	text: '考勤直观图'
		    },
		    title: {
		        text: labTitle
		    },
		    yAxis: {
		        title: {
		            text: '小时（h)'
		        },
		        plotLines: [{
		        	color: 'rgba(124, 181, 236, 1)',
		        	dashStyle: 'solid',
		        	value: workmean,
		        	width: 2
		        },{
		        	color: 'rgba(67, 67, 72, 1)',
		        	dashStyle: 'solid',
		        	value: allmean,
		        	width: 2
		        }]
		    },
		    series: [{
		    	name: '工作日平均时长',
		    	data: allPerpleWorkAvgTime
		    },{
		    	name: '总的平均时长',
		    	data: allPerpleAllAvgTime
		    }]
		};

		$("#chart").highcharts(chartConfig);
    }, 100);
});

		
function isWeekDay(date){
	var tmpDate = new Date(date);
	var day = tmpDate.getDay();
	if(day == '0' || day =='6'){
		return true;
	}else{
		return false;
	}
}










