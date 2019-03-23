import * as jwt from "jsonwebtoken";

export class Session {
    id: string;
    currAuthToken:string;
    tokenPayload:{ [key:string]:any};
    expires:Date;

    constructor(bearer: string) {
        this.id = this.generateId();
        this.currAuthToken = bearer;
        this.tokenPayload = jwt.decode(this.currAuthToken) as {[key:string]:any};
        let payloadObject = this.tokenPayload as {[key:string]:any};
        this.expires = new Date(payloadObject["exp"] * 1000);
    }

    generateId():string {
        return Math.floor(Math.random()*30).toString();
    }
}