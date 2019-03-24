import { Session } from "../Graph/Session";
import { User } from "../Graph/User";
import { AuditLogs } from "../Graph/AuditLogs";
import { MailBox } from "../Graph/MailBox";

export class Tenant {
    users: number;
    mailLeader: string;
    mailsSent:number;
    lastLogin: string;

    constructor(numUsers:number, mailsSent:number, mailLeader:string, lastLogin:string) {
        this.users = numUsers;
        this.mailLeader = mailLeader;
        this.mailsSent = mailsSent;
        this.lastLogin = lastLogin;
    }

    static get(adminSession: Session): Promise<Tenant> {
        return new Promise((resolve, reject) => {
            Promise.all([
                User.getAllUserCount(adminSession.currAuthToken),
                User.getAllUsers(adminSession.currAuthToken),
                AuditLogs.get(adminSession.currAuthToken)
            ]).then(vals => {
                const userCount = vals[0];

                MailBox.getTenantSentMailCountByUser(adminSession.currAuthToken,vals[1]).then(userAndMailCount => {
                    resolve(new Tenant(userCount, userAndMailCount[1],userAndMailCount[0], vals[2]));
                })
            });
        })
    }
}
