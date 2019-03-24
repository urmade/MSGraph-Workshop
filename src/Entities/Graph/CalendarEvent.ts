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
        const options = {
            method: "GET",
            url: "https://graph.microsoft.com/v1.0/me/events",
            headers: {
                Authorization: "bearer " + token
            }
        };
        return new Promise((resolve, reject) => {
            request(options, (error, response, body) => {
                if (error) throw error;
                let eventArr: Array<CalendarEvent> = [];
                JSON.parse(body).value.forEach((rawEvent: { [key: string]: any }) => {
                    eventArr.push(new CalendarEvent(rawEvent));
                });
                resolve(eventArr);
            })
        })
    }
}