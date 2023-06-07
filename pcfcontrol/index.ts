import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as pbi from 'powerbi-client';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import * as Msal from 'msal';

export class pcfcontrol1 implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private powerbi: pbi.service.Service;
   
    
    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
  
    public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): Promise<void> {
        // Add control initialization code
        this.powerbi = new pbi.service.Service(pbi.factories.hpmFactory, pbi.factories.wpmpFactory, pbi.factories.routerFactory);
                  
        const signinContainer = document.createElement("div");
        signinContainer.classList.add("signin-container");
      
        const signinButton = document.createElement("button");
        signinButton.textContent = "Login";
        signinButton.classList.add("btn", "btn-primary");
        signinButton.addEventListener("click", async () => {
          
          const clientId = '9b5f16ff-a73d-43bb-ac7e-b1801f46bf25';
          const clientSecret = 'GDr8Q~6wDGL.ZU~ixqsaCzjHsAZba3~Fgp72dbcM';
                  
          const msalConfig = {
            auth: {
              clientId: clientId,
              authority: 'https://login.microsoftonline.com/09e251dc-5e87-48bf-b4d2-71b01adb984a',
              clientSecret: clientSecret,              
              cacheLocation: "localStorage", 
            }
          };
          const msalInstance = new Msal.UserAgentApplication(msalConfig);

          const loginRequest = {
            scopes: ['https://analysis.windows.net/powerbi/api/.default']
          };
          
          msalInstance.loginPopup(loginRequest).then((response) => {
            console.log('Login successful', response);
          
            msalInstance.acquireTokenSilent(loginRequest).then((response) => {
              console.log('Access token acquired', response.accessToken);
              const accessToken = response.accessToken;

              // Add code to hide the signinContainer and show the report container
              signinContainer.style.display = "none";
              
              //Welcome and logout button
              const welcomecontainer = document.createElement('div');
              welcomecontainer.classList.add('custom-combobox-container', 'control-pane');
              const usernameLabel = document.createElement("label");
              usernameLabel.textContent = "Welcome, "+response.account.name+".";
              usernameLabel.style.fontWeight = "bold";
      
              welcomecontainer.appendChild(usernameLabel);
                          
              
              const comboboxContainer = document.createElement('div');
              comboboxContainer.classList.add('custom-combobox-container', 'control-pane');
              // Create the first combobox
              const combobox1 = document.createElement('select');
              combobox1.classList.add('custom-combobox');
              combobox1.id = 'combobox1';
             // Crie o elemento option com o valor padrão
              const defaultOption = document.createElement('option');
              defaultOption.value = 'default';
              defaultOption.text = 'Select one workspace';
              defaultOption.selected = true;            
              
              // Create the second combobox
              const combobox2 = document.createElement('select');
              combobox2.id = 'combobox2';
              combobox2.classList.add('custom-combobox');
              // Crie o elemento option com o valor padrão
              const defaultOption1 = document.createElement('option');
              defaultOption1.value = 'default';
              defaultOption1.text = 'Select one report';
              defaultOption1.selected = true;
              combobox2.disabled = true;
              // Adicione o elemento option à combobox
              combobox1.appendChild(defaultOption);
              combobox2.appendChild(defaultOption1);
              // Append comboboxes to the combobox container
                            //Create report container
              const reportContainer = document.createElement('div');
              reportContainer.classList.add('custom-report');
              //reportContainer.style.paddingTop = '50px';   
              
              // Append comboboxes to the combobox container
              comboboxContainer.appendChild(combobox1);
              comboboxContainer.appendChild(combobox2);
              
              container.appendChild(welcomecontainer);      
              container.appendChild(comboboxContainer);              
              container.appendChild(reportContainer);
              
              // Add options to combobox1
                    
                // Fetch list of workspaces
                if (combobox1 !== null) {
                  // Fetch list of workspaces
                  fetch('https://api.powerbi.com/v1.0/myorg/groups', {
                    method: 'GET',
                    headers: {
                      'Authorization': `Bearer ${response.accessToken}`
                    }
                  })
                    .then(response => response.json())
                    .then(data => {
                      // Populate workspace select element
                      data.value.forEach((workspace: any) => { // Explicitly define the type of 'workspace' as 'any'
                        const option = document.createElement('option');
                        option.value = workspace.id;
                        option.text = workspace.name;
                        combobox1.appendChild(option);
                      });
                    })
                    .catch(error => {
                      console.error('Error fetching workspaces:', error);
                    });
                } else {
                  console.error('Workspace select element not found');
                }                     

              
              combobox1.addEventListener('change', async () => {
                const selectedWorkspaceId = combobox1.value;
                console.log("Workspaceid: ",selectedWorkspaceId);

                // Fetch reports for the selected workspace
                const reportsResponse = await fetch(`https://api.powerbi.com/v1.0/myorg/groups/${selectedWorkspaceId}/reports`, {
                  method: 'GET',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`
                  }
                });

                const reportsData = await reportsResponse.json();
                const reports = reportsData.value;

                combobox2.innerHTML = '';
                // Adicionar opção personalizada no início da combobox2
                const customOption = document.createElement('option');
                customOption.value = 'custom';
                customOption.text = '';
                combobox2.appendChild(customOption);

                // Add reports to combobox2
                reports.forEach((report: any) => {
                  const option = document.createElement('option');
                  option.value = report.id;
                  option.text = report.name;
                  combobox2.appendChild(option);
                });
                combobox2.disabled = false;
              });             
                    

              combobox2.addEventListener('change', async () => {                
                const reportId = combobox2.value;                        
                const groupId = combobox1.value;
                fetch(`https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}`, {
                  method: 'GET',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`
                  }
                }).then((response) => {
                  return response.json();
                }).then((data) => {
                  console.log("ReportID :",reportId);
                  console.log("WorkspaceID :",groupId);                  
                  const embedUrl = data.embedUrl;
                  console.log("Url embutido: ",embedUrl);
                  // Used the embed token and embed URL to embed the report in the app
                  const config = {
                    type: 'report',
                    accessToken: accessToken,
                    embedUrl: embedUrl,
                    id: reportId,
                    permissions: pbi.models.Permissions.All,
                    settings: {
                      filterPaneEnabled: true,
                      navContentPaneEnabled: true
                    }
                  };
                  
                  this.embedReportInContainer(config, reportContainer);
                  
                  
                }).catch((error) => {
                  console.log('Error generating embed token:', error);
                });  
              });
            
            
  

              ////
            }).catch((error) => {
              console.error('Failed to acquire access token', error);
            });
          }).catch((error) => {
            console.error('Failed to log in', error);
          });
            
        
        }); 
      
        signinContainer.appendChild(signinButton);
        container.appendChild(signinContainer);
      
        // Add Bootstrap CSS
        const bootstrapCSS = document.createElement("link");
        bootstrapCSS.rel = "stylesheet";
        bootstrapCSS.href = "https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css";
        document.head.appendChild(bootstrapCSS);
      
        // Add Bootstrap JavaScript
        const bootstrapJS = document.createElement("script");
        bootstrapJS.src = "https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js";
        document.body.appendChild(bootstrapJS);

        //This will load the powerbi.js script and make the powerbi object available for use.
        const powerbiScript = document.createElement('script');
        powerbiScript.src = 'https://microsoft.github.io/PowerBI-JavaScript/demo/node_modules/powerbi-client/dist/powerbi.js';
        document.head.appendChild(powerbiScript);      
      }

      private async embedReportInContainer(config: pbi.models.IEmbedConfiguration, container: HTMLElement): Promise<void> {
        try {
          // Load the Power BI JavaScript client library
          if (!this.powerbi) {
            this.powerbi = new pbi.service.Service(pbi.factories.hpmFactory, pbi.factories.wpmpFactory, pbi.factories.routerFactory);
          }
      
          // Embed the report using the Power BI JavaScript client
          const report = await this.powerbi.embed(container, config);
      
          // Add event handlers for report events
          report.on('loaded', () => {
            console.log('Report loaded');
          });
      
          report.on('rendered', () => {
            console.log('Report rendered');
          });
      
          report.on('error', (event) => {
            console.error('Report error', event.detail);
          });
      
        } catch (error) {
          console.error('Embedding error', error);
        }
      }                   
           
    

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
        
    }
}
