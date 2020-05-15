---
layout: article
title: Transfer domain host from GoDaddy to Google Domains
caption: Transfer GoDaddy Domain To Google Domains
description: Detailed guide of transferring domain host from GoDaddy to Google Domains
image: transfer-domain-to-google.png
labels: [google domains,godaddy,transfer]
---
This article is a detailed step-by-step guidance to transfer the domain host from GoDaddy to Google Domains. There might be multiple reasons for transferring the domains from one registrar to another. This may include pricing, security, privacy features, flexibility, special offers, hosting options etc.

Some of the benefits of the Google Hosting over the GoDaddy hosting are:

* Free privacy protection for the host
* E-mail encryption for the free e-mail forwarding service

## Unlock account in GoDaddy

The transfer of the account should be initiated from the GoDaddy web-site.

Login to the web-site and click on *Manage* button for the domain you want to transfer.

![Manage domain in GoDaddy](manage-domain.png)

Scroll down to the *Additional Settings* section and turn off the *Domain Lock* feature

![Remove domain lock in the domain](unlock-godaddy-domain.png)

If you have enabled the privacy protection feature with GoDaddy it must be disabled before transferring. It could be reactivated in Google Domains later.

![Removing privacy setting on domain](remove-privacy.png)

It might take several minutes for the feature to be disabled.

![Pending the privacy removing](remove-privacy-pending.png)

> If privacy feature is not turned off, the domain transfer won't complete and the *'Transfer rejected. Check with current registrar for more info'* message will be displayed in the Google Domains page

![Transfer rejected due to the privacy enabled](google-domains-transfer-rejected.png)

Click the *Get authorization code* button to generate temporarily token which needs to be pasted in Google Domain to authorize the transfer. Token will be e-mailed to the e-mail registered with GoDaddy.

![Generating authorization code to unlock domain](get-authorization-code.png){ width=250 }

> It is recommended to generate this token just before the transfer as it can expire and the transfer process may fail.

## Initiating transfer to Google Domains

Login to Google domains and activate the *Transfer* tab and search for the domain you want to transfer (the one you have unlocked in the previous step).

![Transfer domain to google](transfer-domain-to-google.png)

Hit enter on the search bar and follow the wizard to perform the transfer. Make sure that the *[domain name] is unlocked and ready to transfer* message is displayed in the first step. Paste the authorization token e-mailed from the previous step.

![Step 1: Prepare domain](transfer-form.png)

Specify the options you want to transfer from the previous settings (such as DNS records). You might want to use the default selections.

![Step 2: Import web settings](import-web-settings-records.png){ width=350 }

Configure the privacy and auto renew options in the next step

![Step 3: Configure registration settings](config-registry-settings.png){ width=450 }

Fill the purchase form to finalize the transfer.

> Note it is required to pay for one year maintenance of domain ahead to perform the transfer. However the existing period of registration will be preserved. For example, if the domain was due to expire in 1 year (e.g. January 2020) transferring this to Google Domains will extended this for one more year (e.g. January 2021).

![Complete the purchase of a domain](purchase-form.png){ width=450 }

Once purchase is approved domain will be waiting for the approval from GoDaddy to complete the transfer.

![Pending the approval from GoDaddy](pending-domain-waiting-for-approval.png)

## Finalizing transfer

You will receive the e-mail from GoDaddy regarding the transfer which will be automatically completed within several business days. However it is possible to speed up the process and complete the transfer immediately.

To do this navigate to *View details* link in the pending out domain in the *My domains* section in GoDaddy web-site.

![Pending transfers of GoDaddy domains](mydomains-pending-transfer.png)

Click on *Accept or decline* link in the grid.

![Domain transfer details](domain-transfer-details.png)

Select *Accept transfer* option and click *OK*

![Accepting transfer of a domain](accept-transfer.png){ width=450 }

The transfer will be completed in several minutes. Once done domain will be removed from GoDaddy console and will be shown in the Google Domains console.

![Completed transfer](transfer-completed.png)
