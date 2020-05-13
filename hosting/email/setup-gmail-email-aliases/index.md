---
layout: article
title: Setup E-Mail aliases in Gmail to send e-mail from custom domain
caption: Setup E-Mail Aliases In Gmail
description: Guide to setup e-mail aliases of custom domains when sending e-mails from Gmail
lang: en
image: /hosting/email/setup-gmail-email-aliases/smtp-server-details.png
labels: [alias,e-mail,gmail,send as]
---
Gmail enables sending of e-mails with alias, which means that gmail can be sent on behalf of the custom domain setup with registrars such as [GoDaddy](https://godaddy.com) or [Google Domains](https://domains.google).

Below is a detailed step-by-step instruction of setting up the alias for sending e-mail as service without the need of using any additional 3rd party services.

## Setting up application

The first step for enabling the alias is setting up the 2-step verification and app passwords.

Navigate to Google Accounts.

{% include img.html src="google-account.png" width=250 alt="Opening Google account settings" align="center" %}

Select the security tab and enable *2-Step Verification* option. Follow the guide to setup the settings.

{% include img.html src="google-account-security.png" alt="Google account security tab" align="center" %}

Once finished select the *App passwords* option to create new application. Select *Other (custom name)* option from the *select app* drop-down box.

{% include img.html src="create-google-app.png" width=350 alt="Creating new Google App" align="center" %}

Click *Generate* button. As the result the password is generated for the app. Copy this password as it will be used in the next step.

{% include img.html src="generated-app-password.png" width=350 alt="Generated password for the application" align="center" %}

## Setting up e-mail alias

Open the [e-mail page](https://mail.google.com) and select the *Settings* command from the drop-down.

{% include img.html src="google-email-settings.png" height=350 alt="Opening Gmail settings page" align="center" %}

Activate *Accounts and import* tab.

Locate *Send mail as* section and click *Add another e-mail address* link.

{% include img.html src="add-another-email-address.png" alt="Adding another e-mail alias" align="center" %}

Specify the information about the other e-mail (i.e. the one you are creating alias for).

{% include img.html src="email-address-details.png" width=450 alt="Setting up alias details" align="center" %}

Configure the smtp server. Set the value of *SMTP server* to *smtp.gmail.com*. User name is your gmail user name (not the custom domain name, i.e. the currently logged in user to gmail). Password is an app password generated in the previous step (not the gmail password).

{% include img.html src="smtp-server-details.png" width=450 alt="Specifying SMTP settings" align="center" %}

As an optional step set the created alias as default option so e-mails will be sent using the alias by default.

{% include img.html src="send-mail-as-default.png" alt="Setting the default option for send as" align="center" %}

Now you can send e-mails with an alias and it will appear as e-mail with custom domain to the receiver instead of the @gmail.com e-mail.
