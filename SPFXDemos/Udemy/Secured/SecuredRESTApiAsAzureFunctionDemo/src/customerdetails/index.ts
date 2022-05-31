import { 
    Context,
    HttpMethod,
    HttpRequest,
    HttpResponse,
    HttpStatusCode
  } from 'azure-functions-ts-essentials';

  import jwt = require('jsonwebtoken');

const customersDb=require('./customers.json');


const getCustomerById = (customerid: any) => {
    
    const customer = customersDb.find((customer) => {
      return (customer.id === customerid)
    });
  
    return {
      status: HttpStatusCode.OK,
      body: customer
    };
  };

  
const getAllCustomers = () => {
    return {
      status: HttpStatusCode.OK,
      body: customersDb
    };
  };

  const addNewCustomer = (newCustomer) => {
    
    const newCustomers = customersDb;
    newCustomers.push(newCustomer);     
    return {
      status: HttpStatusCode.Created,
      body: newCustomers
    }
  }

  const decodedValidToken = (accessToken: string) => {
    const key: string = '-----BEGIN CERTIFICATE-----\nMIIDBTCCAe2gAwIBAgIQWHw7h/Ysh6hPcXpnrJ0N8DANBgkqhkiG9w0BAQsFADAtMSswKQYDVQQDEyJhY2NvdW50cy5hY2Nlc3Njb250cm9sLndpbmRvd3MubmV0MB4XDTIwMDQyNzAwMDAwMFoXDTI1MDQyNzAwMDAwMFowLTErMCkGA1UEAxMiYWNjb3VudHMuYWNjZXNzY29udHJvbC53aW5kb3dzLm5ldDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALhz3sIYOFgt3i1T5BBZY+0Q7WimFlwiORviz1c7DCjriLu6kEG3srSAOj+h0/f4iEbfMzUL7sOD/b2zY4FAqSOr32RrI5N17glaAf2wCIb7gXEIfXjx9shMEua3kfjaxtT7Ks6G52WbooCgqA5rjm/1A8dQ4lcjQmzAZRBu1M00MC3+TT+h2kR8dNu1ESXmbzwFmO84x5UjriqEv3dclL3mgRSIGaj1iwoOOHJOIL4pOOR7DVVk/c2H0++Hb1EkqzEkfkhxU+x8tV421V6RyRzTQF6T6BqFl07nNAcTLAeHKo3yaqH7RRjhuMd9rxM2pAKyz8QCsBr5L7tI06AMr0kCAwEAAaMhMB8wHQYDVR0OBBYEFOI7M+DDFMlP7Ac3aomPnWo1QL1SMA0GCSqGSIb3DQEBCwUAA4IBAQBv+8rBiDY8sZDBoUDYwFQM74QjqCmgNQfv5B0Vjwg20HinERjQeH24uAWzyhWN9++FmeY4zcRXDY5UNmB0nJz7UGlprA9s7voQ0Lkyiud0DO072RPBg38LmmrqoBsLb3MB9MZ2CGBaHftUHfpdTvrgmXSP0IJn7mCUq27g+hFk7n/MLbN1k8JswEODIgdMRvGqN+mnrPKkviWmcVAZccsWfcmS1pKwXqICTKzd6WmVdz+cL7ZSd9I2X0pY4oRwauoE2bS95vrXljCYgLArI3XB2QcnglDDBRYu3Z3aIJb26PTIyhkVKT7xaXhXl4OgrbmQon9/O61G2dzpjzzBPqNP\n-----END CERTIFICATE-----';
  
  
    return jwt.verify(accessToken, key);
  }

  export async function run(context: Context, req: HttpRequest): Promise<any> {
    let response: any;
    const customerid = req.params
      ? req.params.customerid
      : undefined;

      let blValidRequest: boolean = false;
      let blCustomerReadScope: boolean = false;
      let blCustomerWriteScope: boolean = false;
      let isUser: boolean = false;
      const authorizationHeader: string = req.headers.authorization;

      try {
        const decodedToken = (decodedValidToken(authorizationHeader.replace('Bearer ','')) as any);
        console.log("decoded token is : " + decodedToken);

        const allScopes: string = (decodedToken.scp as string)
       
        blCustomerReadScope = (allScopes.indexOf('Customer.Read') >= 0);
        blCustomerWriteScope = (allScopes.indexOf('Customer.Write') >= 0);
    
       
        isUser = (decodedToken.upn.indexOf('sample@sample.com') !== -1);
    
        blValidRequest = true;
      } 
      catch (err) {
        blValidRequest = false;
        switch (err.name){
          case 'NotBeforeError':
            response = {
              status: HttpStatusCode.Unauthorized,
              body: {
                message: `${err.message} : ${err.date}`
              }
            };
            break;
          case 'TokenExpiredError':
            response = {
              status: HttpStatusCode.Unauthorized,
              body: {
                message: `${err.message} : ${err.expiredAt}`
              }
            };
            break;
          case 'JsonWebTokenError':
            response = {
              status: HttpStatusCode.Unauthorized,
              body: {
                message: `${err.message}`
              }
            };
            break;
          default:
            response = {
              status: HttpStatusCode.Unauthorized,
              body: {
                message: `Error while decoding & validating json web token: ${err.message}`
              }
            };
            break;
        }
      }


      if (blValidRequest) {
    switch (req.method) {
      case 'GET':
        if (blCustomerReadScope) {
        response = customerid ? getCustomerById(customerid) : getAllCustomers();
        }
        else {
          response = {
            status: HttpStatusCode.Unauthorized,
            body: {
              message: 'You must have scope of Customer.Read to get customers details.'
            }
          }
        }
        break;
      case 'POST':
        if (blCustomerWriteScope) {
        response = addNewCustomer(req.body);
        }
        else {
          response = {
            status: HttpStatusCode.Unauthorized,
            body: {
              message: 'You must have Customer.Write scope to add customers.'
            }
          }
        }
        break;
  
      default:
        response = {
          status: HttpStatusCode.BadRequest,
          body: {
            error: {
              type: 'not_supported',
              message: `Currently this method ${req.method} is not supported.`
            }
          }
        };
    }
  }
  
   
    response.headers = {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Credentials': 'true'
    };
  
    context.res = response;
    Promise.resolve();
  }