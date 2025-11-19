import * as React from "react";
import { useState, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { FluentProvider, teamsLightTheme } from "@fluentui/react-components";
import { Card, Avatar, Text, Input, Button, Dialog, DialogSurface } from "@fluentui/react-components";
import { IListDirectoryProps } from './IListDirectoryProps';
import styles from './ListDirectory.module.scss';
const bgImage = require('../assets/bg.png');


interface IEmployee {
  Id: number;
  Employee: { Title: string; EMail: string }; // Person field, with both Title (display name) and EMail
  Title: string;
  Phone: string;
  Location: string;
}

const ListDirectory: React.FC<IListDirectoryProps> = (props) => {
  const [employees, setEmployees] = useState<IEmployee[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [open, setOpen] = useState(false);

  useEffect(() => {
    const getEmployees = async () => {
      try {
        const targetSiteUrl = "{https://{your-tenant}.sharepoint.com/sites/{site-with-list}/}"; 

        const spOtherSite = spfi(targetSiteUrl).using(SPFx(props.context));

        const listName = "Employee List";

        const items: IEmployee[] = await spOtherSite.web.lists
          .getByTitle(listName)
          .items
          .select("Id", "Title", "Phone", "Location", "Employee/Title", "Employee/EMail") 
          .expand("Employee") // here you expand the 'Person' field to access its properties
          .orderBy("Employee", true) // true sorts by name (Title) in ascending order
          .top(250)(); // by default sharepoint only grabs 100 items on the list - this is set to a number for future growth

        console.log("Fetched Employees from another site:", items); // for debugging
        
        setEmployees(items);
      } catch (error) {
        console.error("Error fetching employees from another site: ", error);
      }
    };

    getEmployees();
  }, []);

  const handleSearch = (event: React.ChangeEvent<HTMLInputElement>) => {
    setSearchTerm(event.target.value);
  };

  const showEmployeeLocation = (employeeLocation: string) => {
    // for my usecase only two neighborhoods on 4th floor | all neighborhoods start with unique characters | if it's not pink or blue, it's the 3rd floor
    // from here you can use any logic you want to set the svg file that opens when the 'find employee' button is pressed
    const svgFile = employeeLocation.startsWith("P") || employeeLocation.startsWith("B")
      ? 'https://{your-tenant}.sharepoint.com/sites/{site-with-shared-folder}/floor4.svg'
      : 'https://{your-tenant}.sharepoint.com/sites/{site-with-shared-folder}/floor3.svg';

    fetch(svgFile)
      .then(response => response.text())
      .then(svgText => {
        const container = document.getElementById('svgContainer');
        if (container) {
          container.innerHTML = svgText;

          const svgDoc = container.querySelector('svg');
          if (svgDoc) {
            const seat = svgDoc.getElementById(employeeLocation) as unknown as SVGGraphicsElement;
            if (seat) {
              // make the matching seat is visible
              seat.style.visibility = 'visible';

              // add the pulsing animation to the rect element
              seat.classList.add(styles.pulsingRect);
            }

            // hide all other rect elements except the one with the matching id
            const rects = svgDoc.querySelectorAll('rect');
            rects.forEach(rect => {
              const rectElement = rect as SVGGraphicsElement;
              if (rectElement.id !== employeeLocation) {
                rectElement.style.visibility = 'hidden';  
              }
            });
          }
        }
      })
      .catch(error => console.error('Error loading SVG:', error));

    setOpen(true);
  };

  // update the filter logic to search by either employee name (Title) or email
  const filteredEmployees = employees.filter(employee =>
    employee.Employee.Title.toLowerCase().includes(searchTerm.toLowerCase()) || 
    employee.Employee.EMail.toLowerCase().includes(searchTerm.toLowerCase())   
  );

  return (
    <FluentProvider theme={teamsLightTheme}>
      <div>
        <Input 
          placeholder="Search employee..." 
          onChange={handleSearch} 
          style={{ width: '50%', margin: '15px auto', display: 'block' }} 
        />

        {filteredEmployees.map(employee => {

          const profilePictureUrl = `https://{your-tenant}.sharepoint.com/sites/{site-with-list}/_layouts/15/userphoto.aspx?size=S&accountname=${employee.Employee.EMail}`;

          // set button text and disable state based on employee location
          let buttonText = "Find on 3rd floor";
          let isButtonDisabled = false;

          // logic from earlier to change the text on the button | disables button if the employee is fully remote
          if (employee.Location && (employee.Location.startsWith("B") || employee.Location.startsWith("P"))) {
            buttonText = "Find on 4th floor";
          } else if (employee.Location && employee.Location.startsWith("R")) {
            buttonText = "Fully Remote Employee";
            isButtonDisabled = true;
          }

          return (
            <Card key={employee.Id} style={{ width: '55%', margin: '20px auto', padding:'0', rowGap: '0' }}>
              {/* Header section with background color */}
              <div style={{ backgroundColor: '#005C8F', padding: '10px', borderRadius: '5px 5px 0 0', color: 'white'}}>
                <div style={{ display: 'flex', alignItems: 'center' }}>
                  <Avatar name={employee.Employee.Title} image={{src: profilePictureUrl}}/>
                  <div style={{ marginLeft: '15px' }}>
                    <Text style={{ color: '#C0D731', fontSize: '20px', fontWeight: 'bold' }}>{employee.Employee.Title}</Text>
                  </div> 
                </div>
              </div>
                <div style={{ backgroundImage: `url(${bgImage})`}}><br/>
                <Text style={{fontSize: '16px', marginLeft: '10px', marginBottom: '10px', textShadow: '1px 1px 3px #627580' }}>{employee.Title}</Text><br/>
                <Text style={{fontSize: '16px', marginLeft: '10px', marginBottom: '10px', textShadow: '1px 1px 3px #627580' }}>{employee.Phone}</Text><br/>
              {/* Conditional button text and disabled state */}
              <div style={{ textAlign: 'center'}}>
              <Button 
                onClick={() => showEmployeeLocation(employee.Location)} 
                style={{ margin: '10px auto', width: '80%' }} 
                disabled={isButtonDisabled}  // disable button if the employee is fully remote
              >
                {buttonText}
              </Button>
              </div>
              </div>
            </Card>
          );
        })}

        {/* Dialog to display SVG */}
        <Dialog open={open} onOpenChange={() => setOpen(false)}>
          <DialogSurface style={{ width: '70%', maxWidth: '70%', padding: 0 }}>
            <div id="svgContainer" style={{ width: '100%', height: '100%' }}>
              {/* The SVG will be dynamically loaded here */}
            </div>
          </DialogSurface>
        </Dialog>
      </div>
    </FluentProvider>
  );
};

export default ListDirectory;
