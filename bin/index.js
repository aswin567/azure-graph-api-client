#!/usr/bin/env node

// read in env settings
require('dotenv').config();

const yargs = require('yargs');

const fetch = require('./fetch');
const auth = require('./auth');

const options = yargs
    .usage('Usage: --op <operation_name>')
    .option('op', { alias: 'operation', describe: 'operation name', type: 'string', demandOption: true })
    .argv;

async function main() {
    console.log(`You have selected: ${options.op}`);

    switch (yargs.argv['op']) {
        case 'getUsers':

            try {
                // here we get an access token
                const authResponse = await auth.getToken(auth.tokenRequest);

                // call the web API with the access token
                const users = await fetch.getUsers(authResponse.accessToken);

                // display result
                console.log(users);
            } catch (error) {
                console.log(error);
            }
            break;

        case 'createTeam':
            try {
                // here we get an access token
                const authResponse = await auth.getToken(auth.tokenRequest);

                // call the web API with the access token
                const teams = await fetch.createTeam(authResponse.accessToken);

                // display result
                console.log(teams);
            } catch (error) {
                console.log(error);
            }

            break;
            
        case 'getCalanders':
            try {
                // here we get an access token
                const authResponse = await auth.getToken(auth.tokenRequest);

                // call the web API with the access token
                const teams = await fetch.getCalanders(authResponse.accessToken);

                // display result
                console.log(teams);
            } catch (error) {
                console.log(error);
            }

            break;
        default:
            console.log('Select a Graph operation first');
            break;
    }
};

main();