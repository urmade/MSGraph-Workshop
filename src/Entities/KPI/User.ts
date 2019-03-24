import { CalendarEvent } from "../Graph/CalendarEvent";
import { MailBox } from "../Graph/MailBox";
import { Session } from "../Graph/Session";

export class User {
    thisMonthMeetings: number;
    lastMonthMeetingLength: number;
    mailsSent: number;

    constructor(events: Array<CalendarEvent>, sentMails: number) {
        this.thisMonthMeetings =
            events.filter(
                ev => {
                    return !(ev.start < new Date(new Date().getFullYear(), new Date().getMonth(), 1));
                }
            )
                .length;

        this.lastMonthMeetingLength =
            events.filter(
                ev => {
                    return !(ev.start < new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1));
                }
            )
                .reduce(
                    (acc, ev) => {
                        return acc += new Date(ev.end).getTime() - new Date(ev.start).getTime();
                    }, 0)
            / (60000 * 60);

        this.mailsSent = sentMails;
    }
    static get(session: Session): Promise<User> {
        return new Promise((resolve, reject) => {
            Promise.all([
                CalendarEvent.get(session.currAuthToken),
                MailBox.getSentMailCount(session.currAuthToken, session.tokenPayload.oid)
            ]).then(values => {
                const events = values[0];
                const sentMailCount = values[1];
    
                resolve(new User(events, sentMailCount));
            })
        })
    }
}