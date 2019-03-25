import * as request from "request";
import { kvObj } from "./util";

/**
 * Representation of the most important Attributes of an AAD User. The attributes are not complete.
 */
export class User {
    name: string;
    email: string;
    id: string;
    phone: string;

    constructor(userJson: kvObj) {
        this.name = userJson.displayName || "";
        this.email = userJson.mail || "";
        this.id = userJson.id || "";
        this.phone = userJson.mobilePhone || "";
    }

    /**
     * Method to get the profile data of the user who acquired the token through an JWT that was acquired by that user.
     * @param token JWT token acquired by the user which user data should be received
     */
    static get(token: string): Promise<User> {
        //TODO insert the right HTTP Method as well as the fitting URL to get the basic profile of an user without changing any other parts of the methods. (Hint: The JWT by default contains information about the issuing user, and MS Graph has a way to utilize this property)
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
                if(body.error) console.error(body.error);
                resolve(new User(JSON.parse(body)));
            })
        })
    }
    /**
     * Updates certain fields of an users profile in AAD.
     * @param token JWT token acquired by the user which data should be updated
     * @param id ID of the user who should be updated
     * @param updatedUser JSON representation of the data that should be updated in the user profile
     */
    static update(token: string, id: string, updatedUser: kvObj) {
        //TODO insert the right HTTP Method as well as the fitting URL to update a specific users profile.
        const options = {
            method: "",
            url: "",
            headers: {
                Authorization: token,
                "Content-Type": "application/json"
            },
            json: updatedUser
        };
        request(options, (error, response, body) => {
            if (error) throw error;
            body = JSON.parse(body);
            if(body.error) console.error(body.error);
        })
    }

    /**
     * Function that returns an Array of all users in a tenant (up to 100).
     * @param adminToken Application JWT with sufficient scopes.
     */
    static getAllUsers(adminToken:string): Promise<Array<User>> {
        //TODO insert the right HTTP Method as well as the fitting URL to get all users in one tenant.
        const options = {
            method: "",
            url: "",
            headers: {
                Authorization: "bearer " + adminToken
            }
        };
        return new Promise((resolve,reject) => {
            request(options,(error,response,body) => {
                if(error) throw error;
                let userArr = [];
                body = JSON.parse(body);
                if(body.error) console.error(body.error);
                for(let i = 0; i < body.value.length; i++) {
                    userArr.push(new User(body.value[i]));
                }
                resolve(userArr);
            })
        })
    }

    /**
     * Function to get the number of all users in the system (up to 100 users).
     * @param adminToken Application JWT with sufficient scopes.
     */
    static getAllUserCount(adminToken:string): Promise<number> {
        return new Promise((resolve,reject) => {
            this.getAllUsers(adminToken).then(users => {
                resolve(users.length);
            })
        })
    }
}