import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import * as $ from 'jquery';

import * as strings from 'GaModernSearchApplicationCustomizerStrings';

const LOG_SOURCE: string = 'GaModernSearchApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGaModernSearchApplicationCustomizerProperties {
  // This is an example; replace with your own property
  trackingId: string;
}

declare var ga: any;

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GaModernSearchApplicationCustomizer
  extends BaseApplicationCustomizer<IGaModernSearchApplicationCustomizerProperties> {
    private gaScript: HTMLScriptElement = document.createElement("script");

    private getFreshCurrentPage(): string {
      return window.location.pathname + window.location.search;
    }

    private initGA(){
      if (typeof ga != 'function') {
        this.gaScript.text += `
         (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
         (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
         m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
         })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');`;
      }

      
    var gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${this.properties.trackingId}`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);

    eval(`
          window.dataLayer = window.dataLayer || [];
          function gtag(){dataLayer.push(arguments);}
          gtag('js', new Date());
          gtag('config',  '${this.properties.trackingId}');
        `);

      this.gaScript.text += `
      ga('create', '${this.properties.trackingId}', 'auto');
      ga('require', 'displayfeatures');
      ga('set', 'page', '${this.getFreshCurrentPage()}');
      ga('send', 'pageview');`;
    }
    
    private async sendEventToGa(eventAction: any, eventLabel: any) {
      if (typeof ga === 'function') {
        ga('send', {
          'hitType': 'event',
          'eventCategory': 'Search',
          'eventAction': `${eventAction}`,
          'eventLabel': `${eventLabel}`
        });
      }
      console.log('Send ga: ', eventAction, ': ', eventLabel);
    }

    private getSearchRefiners() {
      var eventAction = "Search Filters";
      $("[class*='filterResultBtn']").on('click', (e) => {
  
        document.body.addEventListener('click', elem => {
          var panel = $("[class*='ms-Panel-scrollableContent']");
          var intervalTimer = setInterval(() => {
            if (panel.css("visibility") != "visible") {
              var filters = $("[class*='linkPanelLayout__selectedFilters'] label");
              filters.each((_, elem) => {
                this.sendEventToGa(eventAction, elem.textContent);
                console.log('Search Filter: ', elem.textContent);
              });
  
              clearInterval(intervalTimer);
              this.getSearchResults();
            }
          });
        });
      });
    }
    
  private getSearchResults() {
    setTimeout(() => {
      this.getSearchRefiners();
      var tags = document.querySelectorAll('.template_contentContainer a');
      tags.forEach((element, i) => {
        setTimeout(() => {
          this.sendEventToGa("Search Result", `${i + 1}: ${element.textContent} - ${element.getAttribute("href")}`);
        }, 500);

        let clickHandler = (ev: MouseEvent) => {
          this.sendEventToGa(`Search Result Clicked`, `${element.textContent} -  ${element.getAttribute("href")}`);
        };

        element.addEventListener("mouseup", clickHandler);
        element.addEventListener("auxclick", clickHandler);
      });
    }, 3000);
  }

    private getSearchTerms() {
      var eventAction = 'Search Terms';
      $("[class*='searchBtn']").on('click', (e) => {
        var searchText = $("[data-sp-feature-tag*='SearchBoxWebPart'] input").val();
        console.error('Click: ', searchText);
        this.sendEventToGa(eventAction, searchText);
        this.getSearchResults();
      });

      $("[data-sp-feature-tag*='SearchBoxWebPart'] input").on('keypress', (e) => {
        let searchText = $("[data-sp-feature-tag*='SearchBoxWebPart'] input").val();
        if (e.keyCode == 13) {
          console.error('Enter: ', searchText);
          this.sendEventToGa(eventAction, searchText);
          this.getSearchResults();
        }
      });

    }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.trackingId;
    if (!message) {
      message = '(Missing Google Analytics Tracking Id.)';
    }

    this.initGA();
    this.getSearchTerms();
    
    this.gaScript.type = "text/javascript";
    document.getElementsByTagName("head")[0].appendChild(this.gaScript);

    return Promise.resolve();
  }
}
