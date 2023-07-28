// Required Modules
const mariadb = require("mariadb");
const XLSX = require('xlsx');

// main function 
async function main() {
   let conn;   
   try {  

      // Connect to your mariadb.
      conn = await mariadb.createConnection({
         host: 'localhost', // Your mariadb's host name or IP address, e.g: localhost.
         user:'root', // Your mariadb's user name, e.g: root.
         password: 'password123', // Your mariadb's user password.
         connectionLimit: 5
      });
      
      // Sql query to retrieve data.  It will return an array of JSON.
      var query_result = await conn.query('SELECT * FROM  mysql.item');
      console.log("query_result", query_result)
      

      // Your data manipulation logic.    
      var  excel_array = query_result.map(v => ({...v, "TOTAL": v.qty * v.price}))
      console.log("excel_array", excel_array)

      // Create Excel after you finished data manipluation.
      createEXCEL(excel_array)

      console.log("end", "end")

   } catch (err) {
      // Manage Errors
      console.log(err);
   } finally {
      // Close Connection
      if (conn) conn.close();
   }
}





// Create Excel function
function createEXCEL(excel_array){

   // Please check sheetjs documentation for more detailed information.
   // https://docs.sheetjs.com/docs/getting-started/examples/export/
   
   const workBook = XLSX.utils.book_new();  
   const workSheet1 = XLSX.utils.json_to_sheet(excel_array);    
   XLSX.utils.book_append_sheet(workBook, workSheet1 , '2023');  
  
   var filename = 'item_2023.xlsx'
   XLSX.writeFile(workBook, filename);
}



// Execute main function.
main();