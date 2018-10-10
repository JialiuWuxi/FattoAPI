const express = require('express');
const router = express.Router();
const graph = require('@microsoft/microsoft-graph-client');


router.get('/', function(req, res, next) {
    res.status('200').send('Empty Get');
});

router.get('/casetype', function(req, res, next) {
    const authToken = getTokenFromHeader(req);

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        getListItems(client, process.env.SITE_NAME, process.env.CASE_CATEGORY_NAME)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode).send(err));

    }else{
        res.status('401').send('empty token');
    }
});

router.get('/branches', function(req, res, next) {
    const authToken = getTokenFromHeader(req);

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        getListItems(client, process.env.SITE_NAME, process.env.BRANCH_LIST_NAME)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode).send(err));

        
    }else{
        res.status('401').send('empty token');
    }

});

router.get('/departments', function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const branchid = req.query.branchid;

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        getListItems(client, process.env.SITE_NAME, process.env.DEPARTMENT_LIST_NAME, 'branchid', branchid)
        .then(result => res.status('200')
        .send(result))
        .catch(err => res.status(err.statusCode).send(err));

    }else{
        res.status('401').send('empty token');
    }

});

router.get('/employees', function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const employeeDepartmentid = req.query.departmentid;

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        getListItems(client, process.env.SITE_NAME, process.env.EMPLOYEE_LIST_NAME, 'departmentid', employeeDepartmentid)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode).send(err));

    }else{
        res.status('401').send('empty token');
    }
});

router.get('/clients', function(req, res, next) {
    const authToken = getTokenFromHeader(req);

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        getListItems(client, process.env.SITE_NAME, process.env.CLIENT_LIST_NAME)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode).send(err));
    }else{
        res.status('401').send('empty token');
    }
});

router.post('/clients', function(req, res, next) {

    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });


        //需要添加重复创建判断
        saveListItems(client, process.env.SITE_NAME, process.env.CLIENT_LIST_NAME, clientInfor)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode || '500').send(err));

    }else{
        res.status('401').send('empty token');
    }
});

router.get('/matters', function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });
        getListItems(client, process.env.SITE_NAME, process.env.MATTER_LIST_NAME)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode).send(err));
    }else{
        res.status('401').send('empty token');


    }       

});

router.post('/matters', function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });    
        //需要添加重复创建判断
        saveListItems(client, process.env.SITE_NAME, process.env.MATTER_LIST_NAME, clientInfor)
        .then(result => res.status('200').send(result))
        .catch(err => res.status(err.statusCode || '500').send(err));

    }else{
        res.status('401').send('empty token');
    }
});

router.get('/me', async function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;    

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });
        const me = (await client
            .api(`/me`)
            .get());
        res.status('200').send(me);

    }else{
        res.status('401').send('empty token');    
    }
});

router.post('/guest', async function(req, res, next) {
    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;    

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });
        const data = {};
        data.invitedUserDisplayName = clientInfor.guestDisplayName;
        data.invitedUserEmailAddress = clientInfor.guestEmailAddress;
        data.inviteRedirectUrl = 'https://m365x937980.sharepoint.com';
        data.sendInvitationMessage = false;


        const resback = await client.api('/invitations').post(data);


        res.status('200').send(resback);
    }else{
        res.status('401').send('empty token')
    }
});

router.post('/groups', async function(req, res, next){
    const authToken = getTokenFromHeader(req);
    const clientInfor = req.body;    

    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });
        const data = {};
        data.displayName = clientInfor.displayName;
        data.mailEnabled = clientInfor.mailEnabled;
        data.mailNickname = clientInfor.mailNickname;
        data.securityEnabled = clientInfor.securityEnabled;


        const resback = await client.api('/groups').post(data);


        res.status('200').send(resback);
    }else{
        res.status('401').send('empty token')
    }
});

router.get('/groups', async function(req, res, next){
    const authToken = getTokenFromHeader(req);
    const query = req.query;
 
    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });
        const groups = (await client
            .api(`/groups`)
            .get());
        res.status('200').send(groups);

    }else{
        res.status('401').send('empty token');    
    }
});




async function getListItems(client, siteName, listName, filterName, filterValue) {
    try{
        //Get root name
        const rootName = (await client
        .api(`/sites/root`)
        .get()).siteCollection.hostname;

        //Get the Site ID
        const siteID = (await client
        .api(`/sites/${rootName}:/sites/${siteName}`)
        .select('id')
        .get()).id
        
        const listID = (await client
        .api(`/sites/${siteID}/lists`)
        .filter(`displayName eq '${listName}'`)
        .get()).value[0].id
        
        if(filterValue && filterName){
            let listItems = await client
            .api(`/sites/${siteID}/lists/${listID}/items`)
            .filter(`fields/${filterName} eq ${filterValue}`)
            .expand('fields')
            .get()
            return listItems;
        }else{
            let listItems = await client
            .api(`/sites/${siteID}/lists/${listID}/items`)
            .expand('fields')
            .get()
            return listItems;
        }
    }catch(err){
        throw err;
    };
}

function getTokenFromHeader(req) {
    let authToken = req.get('Authorization');
    if(authToken){
        let authTokenArray = authToken.split(' ');
        if(authTokenArray[0] == 'Bearer'){
            authToken = authTokenArray[1];
        }else{
            authToken = authTokenArray[0];
        }
        return authToken;
    }else{
        return undefined;
    }
}

async function saveListItems(client, siteName, listName, data) {
    try{
        //Get root name
        const rootName = (await client
        .api(`/sites/root`)
        .get()).siteCollection.hostname;

        //Get the Site ID
        const siteID = (await client
        .api(`/sites/${rootName}:/sites/${siteName}`)
        .select('id')
        .get()).id
        
        const listID = (await client
        .api(`/sites/${siteID}/lists`)
        .filter(`displayName eq '${listName}'`)
        .get()).value[0].id
        

        let newItems = await client
        .api(`/sites/${siteID}/lists/${listID}/items`)
        .post({
            "fields": data
            }
        )

        return newItems;
        
    }catch(err){
        throw err;
    };
}


module.exports = router;