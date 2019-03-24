import * as request from "request";

/**
 * Wrapper for getting the latest login into AAD from the AAD Audit Logs.
 */
export abstract class AuditLogs {
    static get(adminToken:string):Promise<string> {
        //TODO insert the right HTTP Method as well as the fitting URL to get all SignIns from the AAD Audit Logs.
        const options = {
            method: "",
            url: "",
            headers: {
                Authorization: "bearer " + adminToken
            }
        }
        return new Promise((resolve,reject) => {
            request(options, (error,response,body) => {
                if(error) throw error;
                body = JSON.parse(body);
                if(body.error) throw body.error;
                let lastLoginDate = new Date(body.value[0].createdDateTime);
                //Resolve the last login into the system as a human readable string
                resolve(`${body.value[0].userDisplayName}: ${lastLoginDate.getDate()}.${lastLoginDate.getMonth()}.${lastLoginDate.getFullYear()} ${lastLoginDate.getHours()}:${lastLoginDate.getMinutes()}`)
            })
        })
    }
}
