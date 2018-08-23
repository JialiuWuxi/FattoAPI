const express = require('express');
const router = express.Router();
const graph = require('@microsoft/microsoft-graph-client');

router.get('/', function(req, res, next) {
    res.status('200').send();
});

router.get('/casetype', async function(req, res, next) {
    const authToken = getTokenFromHeader(req, res);
    if(authToken){
        const client = graph.Client.init({
            authProvider: (done) => {
              done(null, authToken);
            }
        });

        try{
            //Get root name
            const rootName = (await client
            .api(`/sites/root`)
            .get()).siteCollection.hostname;

            //Get the Site ID
            const siteID = (await client
            .api(`/sites/${rootName}:/sites/${process.env.SITE_NAME}`)
            .select('id')
            .get()).id
            
            const listID = (await client
            .api(`/sites/${siteID}/lists`)
            .filter(`displayName eq 'CaseCategory'`)
            .get()).value[0].id
                
            let listItems = (await client
            .api(`/sites/${siteID}/lists/${listID}/items`)
            .expand('fields')
            .get()).value
            res.status('200').send(JSON.stringify(listItems));
        }catch(err){
            res.status('404').send(err.message);
        };
    }else{
        
    }

    
});




function getTokenFromHeader(req, res) {
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
        res.status('401').send('please login first');
        return undefined;
    }
}

module.exports = router;