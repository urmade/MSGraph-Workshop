import * as request from "request";
import { User } from "./User";

/**
 * Provides methods to receive KPIs extracted from users Exchange Mailboxes.
 */
export abstract class MailBox {

    /**
     * Returns the number of elements in the users "Sent Items" folder by getting all mail folders and reading out the property "totalItemCount".
     * @param token JWT acquired by the user which mailbox should be analyzed
     * @param userId ID of the user which mailbox should be analyzed
     */
    static getSentMailCount(token: string, userId: string): Promise<number> {
        //TODO insert the right HTTP Method as well as the fitting URL to get all mail folders for one specific user without changing any other parts of the methods.
        const options = {
            method: "",
            url: "",
            headers: {
                Authorization: "bearer " + token
            }
        };
        return new Promise((resolve, reject) => {
            request(options, (error, response, body) => {
                if (error) throw error;
                body = JSON.parse(body);
                //Error could happen e.g. because the user hasn't accessed his mailbox yet, meaning it does not exist yet
                if (body.error) {
                    console.error(body.error);
                    resolve(0);
                    return;
                }
                if(body.value && body.value.length > 0) {
                    let sentMails = body.value.filter((mailFolder: { [key: string]: string }) => {
                        //Currently only supports German and English mailboxes, it could be examined if folders share the same ID for every user
                        return mailFolder.displayName == "Gesendete Elemente" || mailFolder.displayName == "Sent Items";
                    })[0].totalItemCount;
                    resolve(sentMails);
                }
                //If the Mailbox is currently set up and no Folders were yet created, return a count of 0 
                else resolve(0);                
            })
        })
    }

    /**
     * Returns the cumulative count of items in the "Sent Items" folders of a given set of users as well as the name of the user with the most sent mails.
     * @param adminToken Application JWT with sufficient permissions.
     * @param users Array of users which mailboxes should be taken into consideration.
     */
    static getTenantSentMailCountByUser(adminToken:string, users: Array<User>): Promise<[string,number]> {
        //Get the count of Sent Mails for every individual user
        let promArr:Array<Promise<number>> = [];
        for(let i = 0; i < users.length; i++) {
            promArr.push(this.getSentMailCount(adminToken, users[i].id));
        }
        return new Promise((resolve,reject) => {
            Promise.all(promArr).then(values => {
                //Gets the index of the highest number in the array, meaning the highest number of sent mails by one user in the tenant.
                let highestUserIndex = values.indexOf(Math.max(...values));
                //As the array is in the same order as the user array (meaning values[i] maps to users[i]), with the index of the highest number in the array the name of the owner of that mailbox can be found. To get the overall sent mails in the tenant, the sent mails of all users accounts are summed up.
                resolve([users[highestUserIndex].name, values.reduce((acc,el) => acc+el)]);
            })
        })
    }
}