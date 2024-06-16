!function(e,t){if("object"==typeof exports&&"object"==typeof module)module.exports=t();else if("function"==typeof define&&define.amd)define([],t);else{var n=t();for(var r in n)("object"==typeof exports?exports:e)[r]=n[r]}}(this,(()=>(()=>{"use strict";var e={178:function(e,t,n){var r=this&&this.__assign||function(){return r=Object.assign||function(e){for(var t,n=1,r=arguments.length;n<r;n++)for(var o in t=arguments[n])Object.prototype.hasOwnProperty.call(t,o)&&(e[o]=t[o]);return e},r.apply(this,arguments)};Object.defineProperty(t,"__esModule",{value:!0}),t.constructSheetColumnMetasFromObject=t.getDefaultMapping=t.getDefaultMappingByValueType=t.getDefaultMappingByType=t.getDirectMapping=t.getDateMapping=t.getBooleanMapping=t.getYesNoMapping=t.getStringMapping=t.getNumberMapping=t.getLookupCodeArrayMapping=t.getLookupCodeMapping=t.isNameEquals=t.yesNoOptions=void 0;var o=n(748),i=-1;function l(e,t){var n,r,o,i;return(null===(r=null===(n=null==e?void 0:e.toLocaleLowerCase)||void 0===n?void 0:n.call(e))||void 0===r?void 0:r.trim())==(null===(i=null===(o=null==t?void 0:t.toLocaleLowerCase)||void 0===o?void 0:o.call(t))||void 0===i?void 0:i.trim())}function a(){return{toExcel:function(e){return null==e||Number.isNaN(e)?null:e},fromExcel:function(e){if(null==e)return null;if("number"==typeof e||"bigint"==typeof e)return e;var t=Number.parseFloat(e);return Number.isNaN(t)?null:t},dataType:"number"}}t.yesNoOptions=[{label:"Yes",value:!0},{label:"No",value:!1}],t.isNameEquals=l,t.getLookupCodeMapping=function(e,t,n){return void 0===t&&(t=null),void 0===n&&(n="number"),{toExcel:function(n){var r,o,i;return null!==(i=null!==(o=null===(r=e.find((function(e){return e.value===n})))||void 0===r?void 0:r.label)&&void 0!==o?o:t)&&void 0!==i?i:"string"==typeof n?n:null},fromExcel:function(t){var n;return(null===(n=e.find((function(e){return l(e.label,t)})))||void 0===n?void 0:n.value)||null},options:e,dataType:n}},t.getLookupCodeArrayMapping=function(e,t,n,r){return void 0===t&&(t=null),void 0===n&&(n=!1),void 0===r&&(r="numbers"),{toExcel:function(n){var r;return null!==(r=null==n?void 0:n.length)&&void 0!==r&&r?n.map((function(n){var r,o=null===(r=e.find((function(e){return e.value===n})))||void 0===r?void 0:r.label;return(null==o?void 0:o.length)?o:t})).filter((function(e){return null!=e})).join(","):""},fromExcel:function(t){return(null==t?void 0:t.length)?t.split(",").filter((function(e){var t;return(null===(t=null==e?void 0:e.trim())||void 0===t?void 0:t.length)>0})).map((function(t){var r,o=null===(r=e.find((function(e){return l(e.label,t)})))||void 0===r?void 0:r.value;if(null==o&&n){var a={label:t,value:i--};return e.push(a),a.value}return o})).filter((function(e){return!!e})):[]},options:e,showOptions:!1,dataType:r}},t.getNumberMapping=a,t.getStringMapping=function(){return{toExcel:function(e){return"".concat(e||"")},fromExcel:function(e){return"".concat(e||"").trim()},dataType:"string"}};var u=["yes","ok","true","allow","visible","y"],c=["no","false","forbidden","disable","hide","invisible","n"];function s(e){return void 0===e&&(e=!1),{toExcel:function(t){return null!=t?!0===t?"Yes":"No":e?"(blank)":null},fromExcel:function(e){var t,n;return!0===e||!1===e?e:null==e||"string"!=typeof e?null:!!u.includes(null===(t=e.toLocaleLowerCase())||void 0===t?void 0:t.trim())||!c.includes(null===(n=e.toLocaleLowerCase())||void 0===n?void 0:n.trim())&&null},options:t.yesNoOptions,dataType:"boolean"}}function f(){return{toExcel:function(e){return e},fromExcel:function(e){try{return e?"number"==typeof e?(t=o.DateTime.fromMillis(Math.round(86400*(e-25569)*1e3))).year>3e3||t.year<1e3?null:t.toFormat("yyyy-MM-dd"):"[object Date]"===Object.prototype.toString.call(e)?(t=o.DateTime.fromJSDate(e,{zone:"utc"})).year>3e3||t.year<1e3?null:t.toFormat("yyyy-MM-dd"):"string"==typeof e?function(e){for(var t=0,n=["yyyy-MM-dd","MM-dd-yyyy","yyyy-M-dd","yyyy.MM.dd","MM.dd.yyyy"];t<n.length;t++){var r=n[t],i=o.DateTime.fromFormat(e,r);if(i&&"Invalid DateTime"!==i.toFormat("yyyy-MM-dd"))return i.toFormat("yyyy-MM-dd")}return null}(e):null:null;var t}catch(e){return null}},dataType:"dateTime"}}function d(e){return void 0===e&&(e="string"),{toExcel:function(e){return e},fromExcel:function(e){return e},dataType:e}}function v(e){switch(e){case"boolean":return s();case"dateTime":return f();case"number":return a();case"string":return{toExcel:function(e){return"".concat(e||"")},fromExcel:function(e){return"".concat(e||"").trim()},dataType:"string"};case"numbers":return{fromExcel:function(e){return(null==e?void 0:e.length)>0?e.split(";"):[]},toExcel:function(e){return e.length>0?e.join(";"):""},dataType:"numbers"};default:return d()}}function p(e){return h(function(e){return"bigint"==typeof e||"number"==typeof e?"number":"boolean"==typeof e?"boolean":null!=e&&"[object Date]"===Object.prototype.toString.call(e)?"dateTime":Array.isArray(e)?"numbers":"string"}(e))}function h(e,t){return null!=e?v(e):p(t)}t.getYesNoMapping=s,t.getBooleanMapping=function(){return{toExcel:function(e){return null!=e?!0===e?"TRUE":"FALSE":null},fromExcel:function(e){var t,n;return!0===e||!1===e?e:"string"!=typeof e||!(null==e?void 0:e.length)||["na","null"].includes(e.toLocaleLowerCase())?null:!!["yes","ok","true","allow","visible"].includes(null===(t=e.toLocaleLowerCase())||void 0===t?void 0:t.trim())||!["no","false","forbidden","disable","mute","hide","invisible"].includes(null===(n=e.toLocaleLowerCase())||void 0===n?void 0:n.trim())},options:t.yesNoOptions,dataType:"boolean"}},t.getDateMapping=f,t.getDirectMapping=d,t.getDefaultMappingByType=v,t.getDefaultMappingByValueType=p,t.getDefaultMapping=h,t.constructSheetColumnMetasFromObject=function(e){var t=[];if(null==e)throw new Error("data can not be null");return Object.entries(e).forEach((function(e){var n=e[0],o=e[1],i=r({fieldName:n,header:n},p(o));t.push(i)})),{columns:t}}},512:function(e,t,n){var r=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(o,i){function l(e){try{u(r.next(e))}catch(e){i(e)}}function a(e){try{u(r.throw(e))}catch(e){i(e)}}function u(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(l,a)}u((r=r.apply(e,t||[])).next())}))},o=this&&this.__generator||function(e,t){var n,r,o,i,l={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:a(0),throw:a(1),return:a(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function a(a){return function(u){return function(a){if(n)throw new TypeError("Generator is already executing.");for(;i&&(i=0,a[0]&&(l=0)),l;)try{if(n=1,r&&(o=2&a[0]?r.return:a[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,a[1])).done)return o;switch(r=0,o&&(a=[2&a[0],o.value]),a[0]){case 0:case 1:o=a;break;case 4:return l.label++,{value:a[1],done:!1};case 5:l.label++,r=a[1],a=[0];continue;case 7:a=l.ops.pop(),l.trys.pop();continue;default:if(!((o=(o=l.trys).length>0&&o[o.length-1])||6!==a[0]&&2!==a[0])){l=0;continue}if(3===a[0]&&(!o||a[1]>o[0]&&a[1]<o[3])){l.label=a[1];break}if(6===a[0]&&l.label<o[1]){l.label=o[1],o=a;break}if(o&&l.label<o[2]){l.label=o[2],l.ops.push(a);break}o[2]&&l.ops.pop(),l.trys.pop();continue}a=t.call(e,l)}catch(e){a=[6,e],r=0}finally{n=o=0}if(5&a[0])throw a[1];return{value:a[0]?a[1]:void 0,done:!0}}([a,u])}}},i=this&&this.__importDefault||function(e){return e&&e.__esModule?e:{default:e}};Object.defineProperty(t,"__esModule",{value:!0}),t.importObjects=t.exportObjects=void 0;var l=n(641),a=n(109),u=i(n(517)),c=n(178);function s(e,t){var n={};return e.eachSheet((function(e,r){var o,i,l;if((null==t?void 0:t.length)&&null!=(l=function(e,t,n){if(e<=n.length){var r=n[e-1];if((0,c.isNameEquals)(r.sheetName,t))return r}var o=n.find((function(e){return(0,c.isNameEquals)(e.sheetName,t)}));return null!=o?o:null==n?void 0:n[e-1]}(r,e.name,t))){var a=l.sheetName,u=l.columns,s=l.showHeader,f=void 0===s||s,d=l.runningInStrictMode,p=void 0===d||d;Object.assign(n,((i={})[a]=v(e,u,f,p),i))}else Object.assign(n,((o={})[e.name]=v(e),o))})),n}function f(e){var t={};return e.eachSheet((function(e,n){Object.assign(t,function(e){var t,n=[],r=1,o=e.name,i=[],l=!1;return e.eachRow((function(e,t){if(1!==t){if(!l&&Array.isArray(e.values))if(e.values[1]!==y){for(var o={},a=1;a<r;a++){var u=e.values[a];"object"==typeof u&&null!=u&&"text"in u&&(u=u.text);var s=i[a-1],f=(0,c.getDefaultMappingByValueType)(u).fromExcel(u);(null==o[s]||"string"==typeof o[s]&&!o[s].length)&&(o[s]=f)}d(o)&&n.push(o)}else l=!0}else{var v=1,p=null;do{(null==(p=e.getCell(v).value)?void 0:p.length)>0&&i.push(p.trim()),r=v,v++}while((null==p?void 0:p.length)>0)}})),(t={})[o]=n,t}(e))})),t}function d(e){return Object.values(e).filter((function(e){return!!e})).length>0}function v(e,t,n,r){void 0===n&&(n=!0),void 0===r&&(r=!0);var o=[],i=!1,l={},a=1;return e.eachRow((function(e,u){if(!i&&Array.isArray(e.values))if(e.values[1]!==y)if(null==t?void 0:t.length){if(1===u){if(1==n&&1==r){for(var s=0;s<t.length;s++)if(!(0,c.isNameEquals)(e.getCell(s+1).value,t[s].header))throw new Error('header with name "'.concat(e.getCell(s+1).value,'" is different with what from meta "').concat(t[s].header,'"'))}else if(!r){if(0==n)throw new Error("we don't support mapping without header in Non Strict Mode. It is easily to make a miss-mapping without header.");s=1;var f=null;do{if((null==(f=e.getCell(s).value)?void 0:f.length)>0){var v=t.find((function(e){return(0,c.isNameEquals)(e.header,f)}));null!=v?l[s]=v:console.warn("".concat(f," Not Founded, we will ignore this column"))}a=s,s++}while((null==f?void 0:f.length)>0)}if(n)return}var p={};if(r)t.forEach((function(t,n){var r=t.fieldName,o=t.fromExcel,i=t.dataType,l=e.values[n+1];"object"==typeof l&&null!=l&&"text"in l&&(l=l.text),p[r]=(null!=o?o:(0,c.getDefaultMapping)(i,l).fromExcel)(l)}));else for(var h=1;h<a;h++){var m=l[h];if(null!=m){var g=m.fieldName,b=m.fromexcel,w=m.dataType,x=e.values[h];"object"==typeof x&&null!=x&&"text"in x&&(x=x.text),p[g]=(null!=b?b:(0,c.getDefaultMapping)(w,x).fromExcel)(x)}}d(p)&&o.push(p)}else o.push(e.values.slice(1).map((function(e){return"object"==typeof e&&null!=e&&"text"in e?e.text:e})));else i=!0})),o}function p(e,t){var n=[],r=e.columns,o=e.rows,i=void 0===o?[]:o,l=e.showHeader,a=void 0===l||l;if(a&&(n.push(r.map((function(e){return e.header||e.fieldName}))),t.columns=r.map((function(e){return{header:e.header||e.fieldName}}))),function(e){var t=e.getCell(h-1,1);t.value=y,t.font={color:{argb:"ffffffff"}}}(t),r.forEach((function(e,n){var r;if(e.required&&(t.getCell(1,n+1).font={color:{argb:"ffff0000"},bold:!0}),e.dataValidations||"boolean"===e.dataType)for(var o=null!==(r=e.dataValidations)&&void 0!==r?r:{options:c.yesNoOptions},i=o.options,l=o.affectedRowCount,u=void 0===l?9999:l,s=i.map((function(e){return e.label})),f=function(e,t,n){for(var r=function(e){for(var t,n="",r=[0,e],o=r[0],i=r[1];i>=0;){o=(t=[i%26,Math.floor(i/26)-1])[0],i=t[1];var l=String.fromCharCode("A".charCodeAt(0)+o);n="".concat(l).concat(n)}return n}(n),o=0;o<t.length;o++){var i=h+o,l=e.getCell(i,n+1);l.font={color:{argb:"ffffffff"}},l.value=t[o]}return"$".concat(r,"$").concat(h,":$").concat(r,"$").concat(h+t.length-1)}(t,s,n),d=a?2:1;d<u;d++)t.getCell(d,n+1).dataValidation={type:"list",allowBlank:!0,formulae:[f],showErrorMessage:!0}})),i.length>0)for(var s=function(e){var o=i[e],l=a?e+2:e+1,s=[];r.forEach((function(e,n){var r=e.fieldName,i=e.toExcel,a=e.dataType,f=u.default.get(o,r);f=i?i(f):(0,c.getDefaultMapping)(a,f).toExcel(f),s.push(f),t.getCell(l,n+1).value=f})),n.push(s)},f=0;f<i.length;f++)s(f);var d=[];r.forEach((function(e,t){var n=e.wch;n&&(d[t]=n)})),function(e,t,n){void 0===n&&(n=[]),(null==e?void 0:e.length)&&e[0].forEach((function(r,o){var i=n[o];if(null==i){var l=Math.max.apply(Math,e.map((function(e){var t,n;return null===(n=null===(t=e[o])||void 0===t?void 0:t.toString())||void 0===n?void 0:n.length})).filter((function(e){return e>0})))+3;t.getColumn(o+1).width=l>15?l:15}else t.getColumn(o+1).width=i}))}(n,t,d)}var h=1e5,y="__BREAKLINE__";function m(e){if(!(null==e?void 0:e.length))throw new Error("input parameter can not be empty");for(var t=Object.assign({},e[0]),n=1;n<e.length;n++)for(var r=0,o=Object.entries(e[n]);r<o.length;r++){var i=o[r],l=i[0],a=i[1];null!=a&&(t[l]=a)}return t}var g=new Function("try {return this===window;}catch(e){ return false;}");function b(e,t){return r(this,void 0,void 0,(function(){var n,r,i;return o(this,(function(o){switch(o.label){case 0:return g()?[4,e.xlsx.writeBuffer()]:[3,2];case 1:return n=o.sent(),r=t.indexOf(".xlsx")>=0?t:"".concat(t).concat(".xlsx"),i=new Blob([n],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),(0,a.saveAs)(i,r),[3,4];case 2:return[4,e.xlsx.writeFile(t)];case 3:o.sent(),o.label=4;case 4:return[2]}}))}))}function w(e,t,n,i){return void 0===n&&(n="Sheet"),void 0===i&&(i=!0),r(this,void 0,void 0,(function(){var r,l;return o(this,(function(o){switch(o.label){case 0:return r=m(e),l=(0,c.constructSheetColumnMetasFromObject)(r).columns,[4,M({fileName:t,sheets:[{columns:l,sheetName:n,showHeader:i,rows:e}]})];case 1:return o.sent(),[2]}}))}))}function x(e,t,n){return void 0===n&&(n=!0),r(this,void 0,void 0,(function(){var r;return o(this,(function(o){switch(o.label){case 0:return r=Object.entries(e).filter((function(e){e[0];var t=e[1];return Array.isArray(t)&&t.length>0})).map((function(e){var t=e[0],r=e[1],o=m(r);return{columns:(0,c.constructSheetColumnMetasFromObject)(o).columns,sheetName:t,showHeader:n,rows:r}})),[4,M({fileName:t,sheets:r})];case 1:return o.sent(),[2]}}))}))}function M(e){return r(this,void 0,void 0,(function(){var t,n,r,i,a,u,c,s;return o(this,(function(o){switch(o.label){case 0:for(t=e.fileName,n=e.sheets,r=new l.Workbook,i=0,a=n;i<a.length;i++)u=a[i],c=u.sheetName,s=r.addWorksheet(c),p(u,s);return[4,b(r,t)];case 1:return o.sent(),[2]}}))}))}function E(e,t,n){return void 0===n&&(n=!1),r(this,void 0,void 0,(function(){var r,i;return o(this,(function(o){switch(o.label){case 0:return r=function(e,t){for(var n={},r=0,o=Object.entries(e);r<o.length;r++){var i=o[r],l=i[0],a=i[1];if(null!=a)n[l]=a;else{if(!t)throw new Error("".concat(l," in parameter is empty, you must provide a value, or set ignoreWarning as true"));console.warn("".concat(l," in parameter is empty, we will skip mapping for this column"))}}return n}(t,n),i=(0,c.constructSheetColumnMetasFromObject)(r).columns,[4,N(e,[{columns:i,sheetName:"",showHeader:!0,runningInStrictMode:!1}])];case 1:return[2,o.sent()]}}))}))}function N(e,t,n){return r(this,void 0,void 0,(function(){var i=this;return o(this,(function(a){switch(a.label){case 0:return null==t||Array.isArray(t)||void 0===n?[3,2]:[4,E(e,t,n)];case 1:return[2,a.sent()];case 2:return null==e?[2,Promise.resolve(null)]:"string"==typeof e?[2,new Promise((function(n,a){return r(i,void 0,void 0,(function(){var r,i;return o(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,(new l.Workbook).xlsx.readFile(e)];case 1:return r=o.sent(),n(null!=t&&Array.isArray(t)&&t.length>0?s(r,t):f(r)),[3,3];case 2:return i=o.sent(),a(i),[3,3];case 3:return[2]}}))}))}))]:[2,new Promise((function(n,a){var u=new FileReader;u.onload=function(e){return r(i,void 0,void 0,(function(){var e,r;return o(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,(new l.Workbook).xlsx.load(u.result)];case 1:return e=o.sent(),n(null!=t&&Array.isArray(t)&&t.length>0?s(e,t):f(e)),[3,3];case 2:return r=o.sent(),a(r),[3,3];case 3:return[2]}}))}))},u.readAsArrayBuffer(e)}))]}}))}))}t.exportObjects=function(e,t,n,i){return void 0===n&&(n=!0),void 0===i&&(i="Sheet"),r(this,void 0,void 0,(function(){return o(this,(function(r){switch(r.label){case 0:return Array.isArray(e)?[4,w(e,t,i,n)]:[3,2];case 1:return r.sent(),[3,6];case 2:return"fileName"in e&&"sheets"in e&&Array.isArray(e.sheets)&&e.sheets.length>0&&null==t?[4,M(e)]:[3,4];case 3:return r.sent(),[3,6];case 4:return[4,x(e,t,n)];case 5:r.sent(),r.label=6;case 6:return[2]}}))}))},t.importObjects=N},939:(e,t)=>{Object.defineProperty(t,"__esModule",{value:!0})},620:function(e,t,n){var r=this&&this.__createBinding||(Object.create?function(e,t,n,r){void 0===r&&(r=n);var o=Object.getOwnPropertyDescriptor(t,n);o&&!("get"in o?!t.__esModule:o.writable||o.configurable)||(o={enumerable:!0,get:function(){return t[n]}}),Object.defineProperty(e,r,o)}:function(e,t,n,r){void 0===r&&(r=n),e[r]=t[n]}),o=this&&this.__exportStar||function(e,t){for(var n in e)"default"===n||Object.prototype.hasOwnProperty.call(t,n)||r(t,e,n)};Object.defineProperty(t,"__esModule",{value:!0}),o(n(939),t),o(n(512),t),o(n(178),t)},641:e=>{e.exports=require("exceljs")},109:e=>{e.exports=require("file-saver")},517:e=>{e.exports=require("lodash")},748:e=>{e.exports=require("luxon")}},t={};return function n(r){var o=t[r];if(void 0!==o)return o.exports;var i=t[r]={exports:{}};return e[r].call(i.exports,i,i.exports,n),i.exports}(620)})()));
//# sourceMappingURL=excel-tool.js.map