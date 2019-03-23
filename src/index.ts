import * as express from "express";
import * as request from "request";
import * as env from "dotenv";
import * as cookies from "cookie-parser";
import { Session } from "./SessionStore/session";

env.config();

const app = express();
app.set("view engine", "ejs");
app.set("views", __dirname + "/../static");

let secret = Math.floor(Math.random() * 20).toString();

app.use(express.static("static"));
app.use(cookies(secret));

let sessions = new Map<string, Session>();

app.get("/", (req, res) => {
    if (req.signedCookies.s && sessions.has(req.signedCookies.s)) {
        let session = sessions.get(req.signedCookies.s);
        if (!checkTokenValidity(session.expires)) {
            res.redirect("/login");
            return;
        }

        getUser(session.currAuthToken).then(user => {
            let userObj: { [key: string]: string } = {
                name: user.displayName,
                email: user.mail,
                id: user.id,
                jobTitle: user.jobTitle,
                phone: user.mobilePhone
            };

            const permissions = checkTokenCapabilites(session.tokenPayload);
            if(permissions.userKPIsEnabled) {
                getCalendarEvents(req.signedCookies.s).then(events => {
                    let userkpi = {
                        thisMonthMeetings: 0,
                        lastMonthMeetingLength: 0,
                        mailsSent: 0
                    }
                    if(events) {
                        userkpi = {
                            thisMonthMeetings: events.filter(ev => {
                                let start = new Date(ev.start.dateTime);
                                return !(start < new Date(new Date().getFullYear(),new Date().getMonth(),1));
                            }).length,
                            lastMonthMeetingLength: events.filter(ev => {
                                let start = new Date(ev.start.dateTime);
                                return !(start < new Date(new Date().getFullYear(),new Date().getMonth()-1,1));
                            }).reduce((acc, ev) => {
                                return acc += new Date(ev.end.dateTime).getTime() - new Date(ev.start.dateTime).getTime();
                            },0),
                            mailsSent: 3
                        }
                    }
                    
                    let tenant = {
                        users: "3",
                        mailLeader: "Spammer",
                        mailsSent: "40"
                    };
                    res.render("index", {
                        editableUser: permissions.editableUser,
                        loggedIn: true,
                        //How many users in tenant, list all groups, get all occupied data in GB
                        tenantEnabled: permissions.tenantEnabled,
                        userKPIsEnabled: permissions.userKPIsEnabled,
                        user: userObj,
                        tenant: tenant,
                        userKPIs: userkpi
                    });
                })
            }
            else res.render("index", {
                editableUser: permissions.editableUser,
                loggedIn: true,
                //How many users in tenant, list all groups, get all occupied data in GB
                tenantEnabled: permissions.tenantEnabled,
                userKPIsEnabled: permissions.userKPIsEnabled,
                user: userObj
            });
        });

    }
    else res.redirect("/login");

})
app.get("/login", (req, res) => {
    //The default scope is always user.read (Get users profile information)
    let scope = "user.read";
    //Check if the login endpoint was called with the usecase query parameter (used to ask for additional scopes)
    if (req.query.usecase) {
        switch (req.query.usecase) {
            case "tenantKPIs": scope = "mail.read.All"; break;
            case "userKPIs": scope = "mail.read calendars.read"; break;
            case "editing": scope = "user.readWrite"; break;
            default: break;
        }
        //With the prompt attribute we can enforce that the user has to give his consent again to all required scopes we're asking for
        scope += "&prompt=consent";
    }

    res.redirect(
        /**
         * [Tenant ID]
         * Can be a 
         * tenant id (e.g. microsoft.onmicrosoft.com or its UUID representation) to enable login for users of that specific organization
         * or
         * "common" to support multi-tenant login (meaning every person with every Microsoft-Account can log into the app (if not specified different by their organization admin))
         * 
         * [Version]
         * Currently v1.0 and v2.0 exist. 
         * Among other differences, v1.0 only allows login from organizational accounts, v2.0 allows login from organizational and personal accounts. 
         * For a full list of differences visit: TODO Link to docs
         */
        "https://login.microsoftonline.com/" + process.env.TENANT_ID + "/oauth2/v2.0/authorize?" +
        //[client_id] Enter the ID of the application registration you want to use
        "client_id=" + process.env.CLIENT_ID +
        /**
         * [response_type]
         * Specifies what Active Directory will return after a successful login action. Must at least be 
         * "code" for the OAuth 2.0 flow (meaning an Authorization_Code to be traded for an Access_token) 
         * and can have 
         * "id_token" for the OpenIdConnect flow (meaning you directly get a JWT token containing basic user information) 
         * and/or 
         * "token" (meaning you get an Access_token right away, this is called the implicit flow and must be manually enabled in die Azure AD App Registration)
         */
        "&response_type=code" +
        //[redirect_uri] The URL which will be called with the registration code as a query parameter 
        "&redirect_uri=" + process.env.REDIRECT_URL +
        /**
         * [response_mode]
         * Specifies with which method the code should be returned to the application. Can be 
         * query (Your callback url will be called by GET <callbackurl>?<response_type>=<value>),
         * form_post (Your callback url will be called by POST <callbackurl> {Content-Type:www-formencoded; body: {<response_type>:<value>}}
         * fragment (???)
         */
        "&response_mode=query" +
        //[state] The state parameter gets passed through the whole authorization workflow and can be used to store information that is important for your own application
        "&state= " +
        /**
         * [scope]
         * The scope parameter specifies the permissions that the token received has to call different Microsoft services
         * In case of Microsoft Graph, it consists of two to three parts: (example: Mail.read.all)
         * The Service to call (eg Mail for Exchange Mails)
         * The Permission on that service (read or readWrite)
         * Optional: The scope of the permission (all or blank for delegated permission)
         */
        "&scope=" + scope
        //Optional parameters
        //[Prompt] Can have the values "login", "consent" and "none" and determines which prompt will be shown to the user when he logs in
        //[Login_hint] Can be used to auto fill-in the email adress or user name of the user who wants to authenticate
        //[Domain_hint] Can has the values "consumers" or "organizations" and determines which type of account can log in
        //[Code_challenge] Used for Security checks done by your app
        //[Code_challenge_method] Used for Security checks done by your app
    );
});

app.get("/api/callback", (req, res) => {
    //The authentification code is passed to the redirect url as a query parameter called 'code'
    const authCode: string = req.query.code;
    //Simple check if the url was called with an authentification code to prevent unnecessary HTTP calls
    if (!authCode) {
        res.status(500).send("There was no authorization code provided in the query. No Bearer token can be requested");
        return;
    }
    //Options for the HTTP call (necessary for the request module call)
    const options = {
        method: "POST",
        //[url] Tenant and version must match the code retrieval parameters
        url: "https://login.microsoftonline.com/" + process.env.TENANT_ID + "/oauth2/v2.0/token",
        form: {
            //[grant_type] Specifies what you want to get back from AAD
            grant_type: "authorization_code",
            //[code] The code we got from the authorize-call which will be traded for a bearer token
            code: authCode,
            //[client_id] Enter the ID of the enterprise application or application registration you want to use
            client_id: process.env.CLIENT_ID,
            //[client_secret] Secret string that is used to prove your app has admin access to the application you want to log in, only necessary for Web Apps
            client_secret: process.env.CLIENT_SECRET,
            //[redirect_uri] The URL which was called with the authorization code (aka usually the current URL)
            redirect_uri: process.env.REDIRECT_URL
        }
    };
    //Execute the HTTP Post call with the options set above
    request(options, function (error, response, body) {
        //Basic check to abort if the call itself couldn't be executed
        if (error) throw new Error(error);
        //Safety check to ensure that if the HTTP call didn't return a JSON (what usually shouldn't happen) the server doesn't crash at trying to parse
        try {
            //Parse JSON string from HTTP response to JSON object
            const json = JSON.parse(body);
            //If Active Directory couldn't hand out a Bearer token it will return a JSON object containing an error description
            if (json.error) res.status(500).send("Error occured: " + json.error + "\n" + json.error_description);
            else {
                //Store the acquired Bearer token in our local runtime to manage user sessions
                let s = new Session(json.access_token);
                sessions.set(s.id, s);
                res.cookie("s", s.id, { signed: true });
                //In this example only the Bearer token will be used. If you requested additional information, e.g. an ID-Token, you may have to send this as well
                res.redirect("/");
            }
        }
        catch (e) {
            res.status(500).send("The token acquirement did not return a JSON. Instead: \n" + body);
        }
    });
});
app.post("/api/user/update", (req, res) => {
    if (sessions.has(req.signedCookies.s)) {
        if(req.query.name) updateUser(req.signedCookies.s, {displayName: req.query.name}) ;
        else if(req.query.phone) updateUser(req.signedCookies.s, {mobilePhone: req.query.phone}) ;
        res.sendStatus(200);
    }
    else res.send(403);
})


function checkTokenCapabilites(tokenPayload: { [key: string]: string }): { editableUser: boolean, loggedIn: boolean, tenantEnabled: boolean, userKPIsEnabled: boolean } {
    let permissionString = tokenPayload.scp.toLowerCase();
    let tokenPermissions = {
        editableUser: false,
        loggedIn: true,
        tenantEnabled: false,
        userKPIsEnabled: false
    }
    if (permissionString.includes("user.readwrite")) tokenPermissions.editableUser = true;
    if (
        permissionString.includes("user.read.all") &&
        permissionString.includes("files.read.all") &&
        permissionString.includes("group.read.all") &&
        permissionString.includes("mail.read") &&
        permissionString.includes("auditlog.read.all")
    ) tokenPermissions.tenantEnabled = true;
    if (
        permissionString.includes("calendars.read") &&
        permissionString.includes("mail.read")
    ) tokenPermissions.userKPIsEnabled = true;

    return tokenPermissions;
}

function checkTokenValidity(tokenExpiration: Date): boolean {
    if (Date.now() - tokenExpiration.getTime() >= 3600000) {
        return false;
    }
    return true;
}
function getUser(token: string): Promise<any> {
    const options = {
        method: "GET",
        url: "https://graph.microsoft.com/v1.0/me",
        headers: {
            Authorization: "bearer " + token
        }
    };
    return new Promise((resolve, reject) => {
        request(options, (error, response, body) => {
            if (error) throw error;
            resolve(JSON.parse(body));
        })
    })
}
function updateUser(sessionId: string, updatedUser: { [key:string]:string }) {
    //https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0
    const session = sessions.get(sessionId);
    if (checkTokenCapabilites(session.tokenPayload).editableUser) {
        const bearer = sessions.get(sessionId).currAuthToken;
        const options = {
            method: "PATCH",
            url: "https://graph.microsoft.com/v1.0/users/"+sessions.get(sessionId).tokenPayload.oid,
            headers: {
                Authorization: bearer,
                "Content-Type": "application/json"
            },
            json: updatedUser
        };
        request(options, (error, response, body) => {
            if (error) throw error;
        })
    }
    else throw "updateUser refused: Additional permissions required!";
}
function getCalendarEvents(sessionId:string): Promise<Array<{[key:string]:any}>> {
    const session = sessions.get(sessionId);
        const options = {
            method: "GET",
            url: "https://graph.microsoft.com/v1.0/me/events",
            headers: {
                Authorization: "bearer " + session.currAuthToken
            }
        };
        return new Promise((resolve, reject) => {
            request(options, (error, response, body) => {
                if (error) throw error;
                resolve(body.value);
            })
        })
}
app.listen(process.env.PORT || 3500);