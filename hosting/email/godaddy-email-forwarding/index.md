---
layout: article
title: Setup GoDaddy e-mail forwarding from custom domain for free
caption: Setup GoDaddy E-Mail Forwarding
description: Setup of up to 100 of free e-mail forwarding from custom domain using GoDaddy
image: /images/codestack-snippet.png
labels: [godaddy,email forwarding]
---
If you have a registered domain with GoDaddy you might want to setup e-mails to be send and received using the custom domain (e.g. info@domain.com).

GoDaddy provides the e-mail hosting service. The plan starts at 5$ per user per months.

GoDaddy also provides a free e-mail forwarding service for up to 100 e-mails. All e-mails sent to the specified e-mail will be redirected to the e-mail of your choice, including free emails (e.g. Gmail, Outlook, Yahoo etc.).

This is a detailed step-by-step guide of setting up e-mail forwarding with GoDaddy.

## Add Forwarding E-Mail

Select the *Manage All* link under the *Workspace Email* section in the GoDaddy console.

> You might need to activate this server by clicking *Redeem* button under the *Additional Products* section on the same page.

{% include img.html src="godaddy-100pack-email-forwarding.png" width=550 alt="Free 100 Pack Email Forwarding" align="center" %}

Click *Create Forward* link in the opened page.

{% include img.html src="create-email-forwarding.png" width=550 alt="Create Forward E-Mail" align="center" %}

Fill the *Forward Email* form. Specify the e-mail you want to forward from (i.e. e-mail with your custom domain). And e-mail you want to forward to (e.g. Gmail).

Specify other options if needed, such as capturing all e-mails sent to your domain.

{% include img.html src="create-forwarding-address.png" width=450 alt="Forward E-Mail details" align="center" %}

## Configure DNS Records

Now it is required to configure the DNS record to enable forwarding.

Click on *DNS* button under the domain.

{% include img.html src="manage-domain-dns.png" alt="Manage domain DNS" align="center" %}

Add DNS records from the following table:

| Type  | Host | Points to                   | Priority | TTL    |
|-------|------|-----------------------------|----------|--------|
| MX    | @    | smtp.secureserver.net       | 0        | 1 Hour |
| MX    | @    | mailstore1.secureserver.net | 10       | 1 Hour |
| CNAME | pop  | pop.secureserver.net        | N/A      | 1 Hour |
| CNAME | imap | imap.secureserver.net       | N/A      | 1 Hour |
| CNAME | smtp | smtpout.secureserver.net    | N/A      | 1 Hour |

{% include img.html src="add-dns-record.png" alt="Add new DNS record" align="center" %}

Validate that records are added correctly by activating the *Tools->Server Settings* menu command. The following dialog should be displayed.

{% include img.html src="dns-records.png" width=350 alt="Validated MX records" align="center" %}

## Receiving E-Mails

Now you can send e-mails from any e-mail address to your newly created e-mail (e.g. info@domain.com). The e-mail will be redirected to the specified e-mail box, while the *to* box will display the e-mail with custom domain.

{% include img.html src="received-email.png" alt="E-mail received via alias" align="center" %}

There is however a limitation with GoDaddy e-mail forwarding as encryption is not supported and the *secureserver.net did not encrypt this message* warning is displayed for all forwarded e-mails:

{% include img.html src="unsecure-email.png" alt="Security warning" align="center" %}

Follow the [Setup Google Domains e-mail forwarding from custom domain for free](/hosting/email/googledomains-email-forwarding/) to setup similar free service with Google Domains which supports e-mails encryption and overcomes this limitation. You will need to [Transfer domain host from GoDaddy to Google Domains](/hosting/domain/transfer-godaddy-domain-to-googledomains/) to use this service.