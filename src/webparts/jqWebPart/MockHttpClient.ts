//import interface ISPList
//You do not need to type the file extension(.ts) when importing from the default module
import { ISPList } from "./JQWebPartWebPart";

//It exports the MockHttpClient class as a default module so that it can be imported in other files
export default class MockHttpClient {
private static _items: ISPList[] = [
{
Header: "Working with SharePoint is fun",
Description: `Mauris mauris ante, blandit et, ultrices a, suscipit eget, quam. Integer ut neque. Vivamus nisi metus, molestie vel, gravida in, condimentum sit
amet, nunc. Nam a nibh. Donec suscipit eros. Nam mi. Proin viverra leo ut
odio. Curabitur malesuada. Vestibulum a velit eu ante scelerisque vulputate.`,
ID: 1
},

{
Header: "Sharepoint development using SPFX",
Description: `Sed non urna. Donec et ante. Phasellus eu ligula. Vestibulum sit amet
purus. Vivamus hendrerit, dolor at aliquet laoreet, mauris turpis porttitor
velit, faucibus interdum tellus libero ac justo. Vivamus non quam. In
suscipit faucibus urna. `,
ID: 2
},
{
Header: "Create jQuery Accordion using SPFX",
Description: `<p>
Nam enim risus, molestie et, porta ac, aliquam ac, risus. Quisque lobortis.
Phasellus pellentesque purus in massa. Aenean in pede. Phasellus ac libero
ac tellus pellentesque semper. Sed ac felis. Sed commodo, magna quis
lacinia ornare, quam ante aliquam nisi, eu iaculis leo purus venenatis dui.
</p>
<ul>
<li>List item one</li>
<li>List item two</li>
<li>List item three</li>
</ul> `,
ID: 3
}
];
// It builds the initial ISPList mock array and returns.
public static get(): Promise<ISPList[]> {
return new Promise<ISPList[]>(resolve => {
resolve(MockHttpClient._items);
});
}
}

// You first need to import the MockHttpClient module to the default webpart module.