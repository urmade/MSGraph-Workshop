import * as request from "request";

/**
 * Representation of the most important attributes of an Exchange Calendar Event as well as a get method to get all Calendar Events for a user.
 */
export class CalendarEvent {
    start: Date;
    end: Date;
    durationInHours: number;

    /**
     * Extracts the most important attributes from an Exchange Calendar Event and parses them into Dates
     * @param jsonObj JSON representation of an Exchange Calendar Event
     */
    constructor(jsonObj: { [key: string]: any }) {
        this.start = new Date(jsonObj.start.dateTime);
        this.end = new Date(jsonObj.end.dateTime);
        this.durationInHours = (this.end.getTime() - this.start.getTime()) / (1000 * 60 * 60);
    }
    /**
     * Gets all Calendar Events of a specific user.
     * @param token JWT acquired by the user which Calendar Events should be read
     */
    static get(token: string): Promise<Array<CalendarEvent>> {
        //TODO insert the right HTTP Method as well as the fitting URL to get all Events for one specific user without changing any other parts of the methods. (Hint: The JWT by default contains information about the issuing user, and MS Graph has a way to utilize this property)
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
                let eventArr: Array<CalendarEvent> = [];
                body.value.forEach((rawEvent: { [key: string]: any }) => {
                    eventArr.push(new CalendarEvent(rawEvent));
                });
                resolve(eventArr);
            })
        })
    }
}