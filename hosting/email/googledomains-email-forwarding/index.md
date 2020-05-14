---
layout: article
title: Setup Google Domains e-mail forwarding from custom domain for free 
caption: Setup Google Domains E-Mail Forwarding
description: Setup of up to 100 of free e-mail forwarding from custom domain using Google Domains
image: /images/codestack-snippet.png
labels: [google domains,e-mail forwarding]
---
Google domains registrar provides a free service to forward up to 100 e-mails from the custom domains. To enable this service navigate to *Email* tab in the Google Domains Console.

Specify the e-mail to forward from the in the first box and e-mail to forward to in the second box.

{% include img.html src="email-forwarding.png" alt="Add new E-Mail forwarding" align="center" %}

verification code will be sent to the destination e-mail. Until the e-mail is verified the forwarding is not enabled:

{% include img.html src="email-verification.png" alt="Pending for e-Mail forwarding verification" align="center" %}

Once e-mail is verified the warning is removed:

{% include img.html src="email-verified.png" alt="Validated E-Mail forwarding record" align="center" %}

Unlike [forwarding with GoDaddy](/hosting/email/godaddy-email-forwarding/), Google Domains enables the encryption for the forwarded e-mails:

{% include img.html src="secure-email.png" alt="Received e-mail with encryption" align="center" %}