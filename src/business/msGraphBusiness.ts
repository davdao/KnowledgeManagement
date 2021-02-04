import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export enum EntityType {
    Message = 'message',
    Event = 'event',
    Drive = 'drive',
    DriveItem = 'driveItem',
    ExternalItem = 'externalItem',
    List = 'list',
    ListItem = 'listItem',
    Site = 'site'
}
export class msGraphBusiness {
    public static async GetCurrentUser(_context: WebPartContext): Promise<any> {
        try {
            _context.msGraphClientFactory.getClient().then((client: MSGraphClient) =>
                client
                .api("/me")
                .get((err, res) => {
                    if(err) {
                        console.error(err);
                        return;
                    }

                    var currentUser = res;
                })
            );
        } catch (error) {
            console.log("");
        }
    }

    public static async GetAllDocuments(_context: WebPartContext): Promise<any> {
        try {

            let postData = 
                {
                    "entityTypes": [
                        "driveItem"
                    ],
                    "query": {
                        "queryString": "*"
                    }
                }
            
            _context.msGraphClientFactory.getClient().then((client: MSGraphClient) =>
                client
                .api("/search/query")
                .top(1)
                .post({ requests: [postData] })
                .then((values => {
                    var doc = values;
                    console.log(doc);
                }))
            );
        } catch (error) {
            console.log("");
        }
    }
}