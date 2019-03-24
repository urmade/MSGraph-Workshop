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
     * Method to get an users profile data through an JWT that was acquired by that user.
     * @param token JWT token acquired by the user which user data should be received
     */
    static get(token: string): Promise<User> {
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
        //https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0
        const options = {
            method: "PATCH",
            url: "https://graph.microsoft.com/v1.0/users/" + id,
            headers: {
                Authorization: token,
                "Content-Type": "application/json"
            },
            json: updatedUser
        };
        request(options, (error, response, body) => {
            if (error) throw error;
        })
    }

    /**
     * Function to get the number of all users in the system (up to 100 users).
     * @param adminToken Application JWT with sufficient scopes.
     */
    static getAllUserCount(adminToken:string): Promise<number> {
        const options = {
            method: "GET",
            url: "https://graph.microsoft.com/v1.0/users?$top=100",
            headers: {
                Authorization: "bearer " + adminToken
            }
        };
        return new Promise((resolve,reject) => {
            request(options,(error,response,body) => {
                resolve(JSON.parse(body).value.length);
            })
        })   
    }

    /**
     * Function that returns an Array of all users in a tenant (up to 100).
     * @param adminToken Application JWT with sufficient scopes.
     */
    static getAllUsers(adminToken:string): Promise<Array<User>> {
        const options = {
            method: "GET",
            url: "https://graph.microsoft.com/v1.0/users?$top=100",
            headers: {
                Authorization: "bearer " + adminToken
            }
        };
        return new Promise((resolve,reject) => {
            request(options,(error,response,body) => {
                let userArr = [];
                body = JSON.parse(body);
                for(let i = 0; i < body.value.length; i++) {
                    userArr.push(new User(body.value[i]));
                }
                resolve(userArr);
            })
        })
    }
}