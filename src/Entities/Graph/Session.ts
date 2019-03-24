import * as jwt from "jsonwebtoken";
import { Capabilities } from "../KPI/Capabilities";
import * as request from "request";

/**
 * A Session consists of a JWT Token and the session id for the user / application id who acquired the token as well as important payload data of the JWT that was extracted for easier accessibility.
 */
export class Session {
    id: string;
    currAuthToken:string;
    //TODO Implement tokenPayLoad
    tokenPayload:AADjwt;
    expires:Date;

    //An ID can be provided so that applications can give their ID as an identifier of the session (as this ID doesn't leave the Server, it doesn't have to be hidden like the user id)
    constructor(bearer: string, id?:string) {
        //For user tokens, a session ID is generated that is sent to the client as an encrypted cookie
        this.id = id || this.generateId();
        this.currAuthToken = bearer;
        this.tokenPayload = new AADjwt(this.currAuthToken);
        this.expires = this.tokenPayload.exp;
    }

    generateId():string {
        return Math.floor(Math.random()*30).toString();
    }

    /**
     * This function checks for every action that would be possible with our application (Edit phone number, get Tenant KPIs, get User KPIs), if the session token contains sufficient permissions to execute the necessary operations
     */
    checkCapabilities():Capabilities {
        let permissions = new Capabilities();
        permissions.loggedIn = true;
        if (this.tokenPayload.scp.includes("user.readwrite")) permissions.editableUser = true;
        if (
            this.tokenPayload.scp.includes("user.read.all") &&
            this.tokenPayload.scp.includes("files.read.all") &&
            this.tokenPayload.scp.includes("group.read.all") &&
            this.tokenPayload.scp.includes("mail.read") &&
            this.tokenPayload.scp.includes("auditlog.read.all")
        ) permissions.tenantEnabled = true;
        if (
            this.tokenPayload.scp.includes("calendars.read") &&
            this.tokenPayload.scp.includes("mail.read")
        ) permissions.userKPIsEnabled = true;

        return permissions;
    }

    /**
     * Checks if the token is still valid
     */
    checkValidity():boolean {
        if (Date.now() - this.expires.getTime() >= 3600000) {
            return false;
        }
        return true;
    }

    //This function will only work in the Demo tenant as you have to provide a discrete tenant to get an App JWT. Change the tenant identifier to enable this to work in a tenant of your choice.
    /**
     * Acquires a new JWT in the name of the Application with the scope permissions set in the Azure Active Directory Portal and returns an Application session with the new JWT and the ID of the Application.
     */
    static logInAsApplication():Promise<Session> {
        const options = {
            method: "POST",
            url: "https://login.microsoftonline.com/urmade.onmicrosoft.com/oauth2/v2.0/token",
            form: {
                grant_type: "client_credentials",
                client_id: process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                scope: "https://graph.microsoft.com/.default"
            }
        };
        return new Promise((resolve,reject) => {
            request(options, function (error, response, body) {
                if (error) throw error;
                try {
                    const json = JSON.parse(body);
                    if (json.error) throw(json.error);
                    else {
                        resolve(new Session(json.access_token,process.env.CLIENT_ID));
                    }
                }
                catch (e) {
                    throw("The application token acquirement did not return a JSON. Instead: \n" + body);
                }
            });
        })
    }
}

/**
 * Reprenstation of an Azure Active Directory scheduled JWT. Not all payload data is represented in this Object and the scope is unified as it has different names for App and User tokens.
 */
class AADjwt {
    exp:Date;
    app_displayname:string;
    appid:string;
    ipaddr:string;
    family_name:string;
    given_name:string;
    name:string;
    oid:string;
    scp:Array<string>;
    unique_name:string;

    constructor(jwtString:string) {
        let payload = jwt.decode(jwtString) as {[key:string]:any};

        this.exp = new Date(payload.exp*1000);
        this.app_displayname = payload.app_displayname;
        this.appid = payload.appid;
        this.ipaddr = payload.ipaddr;
        this.family_name = payload.family_name;
        this.given_name = payload.given_name;
        this.name = payload.name;
        this.oid = payload.oid;
        this.unique_name = payload.unique_name;

        if(payload.scp) this.scp = payload.scp.toLowerCase().split(" ");
        else if (payload.roles) {
            this.scp = [];
            for(let i = 0; i < payload.roles.length; i++) {
                this.scp.push(payload.roles[i].toLowerCase());
            }
        }
        else this.scp = [];
    }
}