console.log("run protocol handling");
navigator.registerProtocolHandler("mailto",
                                  "https://outlook.office.com/?path=/mail/action/compose&to=%s",
                                  "Mailto: Links");
