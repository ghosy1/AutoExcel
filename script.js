const excel_file = document.getElementById('excel_file');
const arrange = document.getElementById('arrange')
const newList = document.getElementById('newList')

excel_file.addEventListener('change', (event) => {

    if(!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type))
    {
        document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';

        excel_file.value = '';

        return false;
    }

    var reader = new FileReader();

    reader.readAsArrayBuffer(event.target.files[0]);

    reader.onload = function(event){

        var data = new Uint8Array(reader.result);

        var work_book = XLSX.read(data, {type:'array'});

        var sheet_name = work_book.SheetNames;

        var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {header:1});

        if(sheet_data.length > 0)
        {
            var table_output = '<table class="table table-striped table-bordered">';

            for(var row = 0; row < sheet_data.length; row++)
            {

                table_output += '<tr>';

                for(var cell = 0; cell < sheet_data[row].length; cell++)
                {

                    if(row == 0)
                    {

                        table_output += '<th>'+sheet_data[row][cell]+'</th>';

                    }
                    else
                    {

                        table_output += '<td>'+sheet_data[row][cell]+'</td>';

                    }

                }

                table_output += '</tr>';

            }

            table_output += '</table>';

            document.getElementById('excel_data').innerHTML = table_output;
        }

        excel_file.value = '';

    }

});

function ExportToExcel(type, fn, dl) {
    var elt = document.getElementById('tbl_exporttable_to_xls');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
        XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }) :
        XLSX.writeFile(wb, fn || ('MySheetName.' + (type || 'xlsx')));
}

arrange.addEventListener("click", ()=>{
    var rows = document.querySelectorAll('tr')
    
    
     for (var n=2; n<rows.length; n++)
     {  
        var cell = rows[n].querySelectorAll('td')
        var room = cell[1].textContent
        var name = cell[2].textContent
        var gender = cell[3].textContent
        if (gender == "F"){gender = "Ž"}
        var country = cell[4].textContent
        country = country.toLowerCase()
        country = country.charAt(0).toUpperCase() + country.slice(1)
        var age = cell[5].textContent
        age = age.substring(0, age.indexOf(","));
        var passport = cell[6].textContent
        var firstTwo = passport.slice(0,2)
        if (firstTwo == "2 " ) {passport = "P:" + passport.slice(2)}
        if (firstTwo == "27" ) {passport = "LK:" + passport.slice(2)}
        if (firstTwo == "6 " ) {passport = "DP:" + passport.slice(2)}
        if (firstTwo == "32" ) {passport = "VD:" + passport.slice(2)}
        

        if (gender != "Spol"){
        var newListRow = document.createElement("tr")
        newListRow.classList.add("list")
        newListRow.setAttribute('id',room);
        newList.appendChild(newListRow)

        var newListCell = document.createElement("td")
        newListCell.style.padding="3px"
        newListRow.appendChild(newListCell)
        newListCell.textContent= name

        var newListCell = document.createElement("td")
        newListCell.style.padding="3px"
        newListRow.appendChild(newListCell)
        newListCell.textContent= country

        var newListCell = document.createElement("td")
        newListCell.style.padding="3px"
        newListRow.appendChild(newListCell)
        newListCell.textContent= passport

        var newListCell = document.createElement("td")
        newListCell.style.padding="3px"
        newListRow.appendChild(newListCell)
        newListCell.textContent= age

        
        var newListCell = document.createElement("td")
        newListCell.style.padding="3px"
        newListRow.appendChild(newListCell)
        newListCell.textContent= gender

       }
        
        

        
      }
    var list = document.querySelectorAll(".list")

        var arr = []
        i=0
        list.forEach(listItem=>{
            arr.push(listItem.outerHTML)
            }
        )

        arr.sort()
        list.forEach(element => {
            element.innerHTML=arr[i]
            i++
            
        });
        console.log(arr);
})