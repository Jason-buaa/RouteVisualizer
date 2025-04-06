/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

    /* global console, document, Excel, Office */
    var coordtransform=require('coordtransform');
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
        document.getElementById("createTable").onclick = createTable; // Add this line
      }
    });
       
    
    export async function createTable() {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const table = sheet.tables.add("A1:D1", true);
          table.name = "GPSData";
          table.getHeaderRowRange().values =
            [["Longitude", "Latitude", "NewLongitude", "NewLatitude"]];
          table.rows.add(null, [[118, 32, 0, 0]]);
          await context.sync();
          console.log("Table created successfully.");
        });
      } catch (error) {
        console.error(error);
      }
    }
    
    export async function run() {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const table = sheet.tables.getItem("GPSData");
          // Load the table data
          table.load("columns/items");
          await context.sync();
          const longitudeColumn = table.columns.getItemAt(0).getDataBodyRange();
          const latitudeColumn = table.columns.getItemAt(1).getDataBodyRange();
          const newLongitudeColumn = table.columns.getItemAt(2).getDataBodyRange();
          const newLatitudeColumn = table.columns.getItemAt(3).getDataBodyRange();
          longitudeColumn.load("values");
          latitudeColumn.load("values");
          await context.sync();
          const longitudes = longitudeColumn.values;
          const latitudes = latitudeColumn.values;
          const newLongitudes = [];
          const newLatitudes = [];
          for (let i = 0; i < longitudes.length; i++) {
            const [newLongitude, newLatitude] = coordtransform.wgs84togcj02(longitudes[i][0], latitudes[i][0]);
            newLongitudes.push([newLongitude]);
            newLatitudes.push([newLatitude]);
          }
          newLongitudeColumn.values = newLongitudes;
          newLatitudeColumn.values = newLatitudes;
          await context.sync();
          console.log("Coordinates converted and updated successfully.");
        });
      } catch (error) {
        console.error(error);
      }
    }
    // JSAPI2.0 使用覆盖物动画必须先加载动画插件
    AMap.plugin('AMap.MoveAnimation', function(){
        var marker, lineArr = [[118.005531329652,31.9980646616492],
        [118.00874144761,32.0014735896886],
    ]        
        var map = new AMap.Map("container", {
            resizeEnable: true,
            center: [116.397428, 39.90923],
            zoom: 17
        });

        marker = new AMap.Marker({
            map: map,
            position: [116.478935,39.997761],
            icon: "https://a.amap.com/jsapi_demos/static/demo-center-v2/car.png",
            offset: new AMap.Pixel(-13, -26),
        });

        // 绘制轨迹
        var polyline = new AMap.Polyline({
            map: map,
            path: lineArr,
            showDir:true,
            strokeColor: "#28F",  //线颜色
            // strokeOpacity: 1,     //线透明度
            strokeWeight: 6,      //线宽
            // strokeStyle: "solid"  //线样式
        });
        var passedPolyline = new AMap.Polyline({
            map: map,
            strokeColor: "#AF5",  //线颜色
            strokeWeight: 6,      //线宽
        });


        marker.on('moving', function (e) {
            passedPolyline.setPath(e.passedPath);
            map.setCenter(e.target.getPosition(),true)
        });

        map.setFitView();

        window.startAnimation = async function startAnimation() {
          try {
            await Excel.run(async (context) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const table = sheet.tables.getItem("经纬度");
        
              // 加载表格数据
              table.load("columns/items");
              await context.sync();
        
              const longitudeColumn = table.columns.getItemAt(0).getDataBodyRange();
              const latitudeColumn = table.columns.getItemAt(1).getDataBodyRange();
              longitudeColumn.load("values");
              latitudeColumn.load("values");
              await context.sync();
        
              const longitudes = longitudeColumn.values;
              const latitudes = latitudeColumn.values;
        
              // 构建新的路径数组
              const lineArr = [];
              for (let i = 0; i < longitudes.length; i++) {
                lineArr.push([longitudes[i][0], latitudes[i][0]]);
              }
        
              if (lineArr.length === 0) {
                console.error("表格 '经纬度' 中没有数据");
                return;
              }
        
              // 更新 marker 的路径并开始动画
              marker.moveAlong(lineArr, {
                // 每一段的时长
                duration: 500, // 可根据实际采集时间间隔设置
                // JSAPI2.0 是否延道路自动设置角度在 moveAlong 里设置
                autoRotation: true,
              });
        
              console.log("动画已开始，路径已更新");
            });
          } catch (error) {
            console.error("读取表格或更新路径时出错:", error);
          }
        };

        window.pauseAnimation = function () {
            marker.pauseMove();
        };

        window.resumeAnimation = function () {
            marker.resumeMove();
        };

        window.stopAnimation = function () {
            marker.stopMove();
        };
    });

