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

        getListItems(client, process.env.SITE_NAME, process.env.CASE_CATEGORY_NAME).then(result => {
            if(result) {
                res.status('200').send(result);
            }else{
                res.status('400').send('not recourse');
            }
        })

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

        getListItems(client, process.env.SITE_NAME, process.env.BRANCH_LIST_NAME).then(result => {
            if(result) {
                res.status('200').send(result);
            }else{
                res.status('400').send('not recourse');
            }
        })

        
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

        getListItems(client, process.env.SITE_NAME, process.env.DEPARTMENT_LIST_NAME, 'branchid', branchid).then(result => {
            if(result) {
                res.status('200').send(result);
            }else{
                res.status('400').send('not recourse');
            }
        });

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
        .catch(res.status('400'.send(err)));

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

module.exports = router;