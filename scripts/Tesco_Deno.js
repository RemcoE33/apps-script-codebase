//Importing dependencies
import { listenAndServe } from "https://deno.land/std/http/mod.ts";

//Create server that listens to incomming requests
listenAndServe({ port: 8000 }, async (req) => {
  const { urls } = await req.json();
  const output = [];

  for (let i = 0; i < urls.length; i++) {
    const request = await fetch(urls[i]);
    const html = await request.text();
    const data = JSON.parse(
      /<script type="application\/ld\+json">(.*?)<\/script>/gims.exec(html)[1]
    )[2];

    output.push(data);
  }

  //Returning data to apps script function.
  return new Response(JSON.stringify(output), {
    headers: {
      "content-type": "application/json",
    },
  });
});
