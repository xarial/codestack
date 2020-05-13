---
layout: article
title: Transfer domain host from GoDaddy to Google Domains
caption: Transfer GoDaddy Domain To Google Domains
description: Detailed guide of transferring domain host from GoDaddy to Google Domains
lang: en
image: /hosting/domain/transfer-godaddy-domain-to-googledomains/transfer-domain-to-google.png
labels: [google domains,godaddy,transfer]
---
This article is a detailed step-by-step guidance to transfer the domain host from GoDaddy to Google Domains. There might be multiple reasons for transferring the domains from one registrar to another. This may include pricing, security, privacy features, flexibility, special offers, hosting options etc.

Some of the benefits of the Google Hosting over the GoDaddy hosting are:

* Free privacy protection for the host
* E-mail encryption for the free e-mail forwarding service

## Unlock account in GoDaddy

The transfer of the account should be initiated from the GoDaddy web-site.

Login to the web-site and click on *Manage* button for the domain you want to transfer.

{% include img.html src="manage-domain.png" alt="Manage domain in GoDaddy" align="center" %}

Scroll down to the *Additional Settings* section and turn off the *Domain Lock* feature

{% include img.html src="unlock-godaddy-domain.png" alt="Remove domain lock in the domain" align="center" %}

If you have enabled the privacy protection feature with GoDaddy it must be disabled before transferring. It could be reactivated in Google Domains later.

{% include img.html src="remove-privacy.png" alt="Removing privacy setting on domain" align="center" %}

It might take several minutes for the feature to be disabled.

{% include img.html src="remove-privacy-pending.png" alt="Pending the privacy removing" align="center" %}

> If privacy feature is not turned off, the domain transfer won't complete and the *'Transfer rejected. Check with current registrar for more info'* message will be displayed in the Google Domains page

{% include img.html src="google-domains-transfer-rejected.png" alt="Transfer rejected due to the privacy enabled" align="center" %}

Click the *Get authorization code* button to generate temporarily token which needs to be pasted in Google Domain to authorize the transfer. Token will be e-mailed to the e-mail registered with GoDaddy.

{% include img.html src="get-authorization-code.png" width=250 alt="Generating authorization code to unlock domain" align="center" %}

> It is recommended to generate this token just before the transfer as it can expire and the transfer process may fail.

## Initiating transfer to Google Domains

Login to Google domains and activate the *Transfer* tab and search for the domain you want to transfer (the one you have unlocked in the previous step).

{% include img.html src="transfer-domain-to-google.png" alt="Transfer domain to google" align="center" %}

Hit enter on the search bar and follow the wizard to perform the transfer. Make sure that the *[domain name] is unlocked and ready to transfer* message is displayed in the first step. Paste the authorization token e-mailed from the previous step.

{% include img.html src="transfer-form.png" alt="Step 1: Prepare domain" align="center" %}

Specify the options you want to transfer from the previous settings (such as DNS records). You might want to use the default selections.

{% include img.html src="import-web-settings-records.png" width=350 alt="Step 2: Import web settings" align="center" %}

Configure the privacy and auto renew options in the next step

{% include img.html src="config-registry-settings.png" width=450 alt="Step 3: Configure registration settings" align="center" %}

Fill the purchase form to finalize the transfer.

> Note it is required to pay for one year maintenance of domain ahead to perform the transfer. However the existing period of registration will be preserved. For example, if the domain was due to expire in 1 year (e.g. January 2020) transferring this to Google Domains will extended this for one more year (e.g. January 2021).

{% include img.html src="purchase-form.png" width=450 alt="Complete the purchase of a domain" align="center" %}

Once purchase is approved domain will be waiting for the approval from GoDaddy to complete the transfer.

{% include img.html src="pending-domain-waiting-for-approval.png" alt="Pending the approval from GoDaddy" align="center" %}

## Finalizing transfer

You will receive the e-mail from GoDaddy regarding the transfer which will be automatically completed within several business days. However it is possible to speed up the process and complete the transfer immediately.

To do this navigate to *View details* link in the pending out domain in the *My domains* section in GoDaddy web-site.

{% include img.html src="mydomains-pending-transfer.png" alt="Pending transfers of GoDaddy domains" align="center" %}

Click on *Accept or decline* link in the grid.

{% include img.html src="domain-transfer-details.png" alt="Domain transfer details" align="center" %}

Select *Accept transfer* option and click *OK*

{% include img.html src="accept-transfer.png" width=450 alt="Accepting transfer of a domain" align="center" %}

The transfer will be completed in several minutes. Once done domain will be removed from GoDaddy console and will be shown in the Google Domains console.

{% include img.html src="transfer-completed.png" alt="Completed transfer" align="center" %}
