import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'RestrictExtensionApplicationCustomizerStrings';
import "../../ExternalRef/css/alertify.min.css";
import "../../ExternalRef/css/style.css";
import pnp from 'sp-pnp-js';
import * as $  from 'jquery'
import "alertifyjs";

var alertify: any = require("../../ExternalRef/js/alertify.min.js");
const LOG_SOURCE: string = 'RestrictExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
var restrictUrlArray=[];
var domainArray=[];

export interface IRestrictExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class RestrictExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IRestrictExtensionApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
 
    return super.onInit().then(() => {  
      pnp.setup({
         spfxContext: this.context
         });
         this.getListItems();
        //  return Promise.resolve();
    }); 
   
    
  }

 async getListItems()
  {
    var currentUserName=this.context.pageContext.user.email;
    var homepageURL=this.context.pageContext.web.absoluteUrl;
    var splittedarray=currentUserName.split("@");
    var lowerEmail=splittedarray[1].toLowerCase();
    var emailIdx=lowerEmail.indexOf('browardbehavioralhc');
    var SecemailIdx=lowerEmail.indexOf('cariskpartners');

    await pnp.sp.web.lists.getByTitle('DomainList').items.select("Title").get().then((DomainItems: any[]) => {
      // console.log(DomainItems);
      if(DomainItems.length>0)
      {
        for (var index = 0; index < DomainItems.length; index++) {
          var domainString=DomainItems[index].Title;
          domainArray.push(domainString);
        }
        // console.log(domainArray);
      }

    });
if(domainArray.length>0)
{
  let resultDomain = domainArray.filter(function (Domainvalue) {
    var toLowerDomain=Domainvalue.toLowerCase();
    var domainIndex=lowerEmail.indexOf(toLowerDomain);
    if(domainIndex>=0)
    {
      return Domainvalue;
    }
});


if(resultDomain.length<=0)
{
await pnp.sp.web.lists.getByTitle('RestrictUrlList').items.select("Title").get().then((allItems: any[]) => {
  // console.log(allItems)
  if(allItems.length>0)
  {
    for (var index = 0; index < allItems.length; index++) {
      var splitString=allItems[index].Title.split('?');
      restrictUrlArray.push(splitString[0]);
    }
    // console.log(restrictUrlArray);
  }

});
if(restrictUrlArray.length>0)
{
  var locationURL=window.location.href.toLowerCase().split('?');
  var splittednewLoc=locationURL[0];
  let result = restrictUrlArray.filter(function (urlvalue) {
    var toLower=urlvalue.toLowerCase();
    return splittednewLoc==toLower;
});
// console.log(result);
if(result.length>0)
{
  let message: string = "Sorry! You are not authorized to access this page";
  alertify.alert(message, function() {
      window.location.href=homepageURL;
  }).set({ 'closable':false})
  .setHeader("<em> Alert </em> ");
}
}


}
}



  }
}
