export class Capabilities {
    editableUser: boolean;
    loggedIn: boolean;
    tenantEnabled: boolean;
    userKPIsEnabled: boolean;

    constructor() {
        this.editableUser = false;
        this.loggedIn = false;
        this.tenantEnabled = false;
        this.userKPIsEnabled = false;
    }
}