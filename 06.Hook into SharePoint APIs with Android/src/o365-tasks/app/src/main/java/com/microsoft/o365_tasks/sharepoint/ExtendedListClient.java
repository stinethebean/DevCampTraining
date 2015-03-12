package com.microsoft.o365_tasks.sharepoint;

import android.util.Log;
import android.util.Xml;

import com.google.common.base.Charsets;
import com.google.common.util.concurrent.FutureCallback;
import com.google.common.util.concurrent.Futures;
import com.google.common.util.concurrent.ListenableFuture;
import com.google.common.util.concurrent.SettableFuture;
import com.microsoft.sharepointservices.Credentials;
import com.microsoft.sharepointservices.ListClient;
import com.microsoft.sharepointservices.SPList;

import org.apache.http.util.EncodingUtils;
import org.json.JSONObject;
import org.xmlpull.v1.XmlSerializer;

import java.io.StringWriter;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Map;

/**
 * This class adds a simple "createList" function to the base ListClient class, which allows us
 * to create a List on a SharePoint site based on a template.
 *
 * This class is only required for the purposes of this specific demo and is not part of the Office 365 SDK for Android
 */
public class ExtendedListClient extends ListClient {

    private static final String TAG = "ExtendedListClient";
    
    public ExtendedListClient(String serverUrl, String siteRelativeUrl, Credentials credentials) {
        super(serverUrl, siteRelativeUrl, credentials);
    }
    
    public ListenableFuture<SPList> createList(String listName, String listTemplate) {
        final SettableFuture<SPList> result = SettableFuture.create();
        
        final String createListUrl = getSiteUrl() + "_api/web/lists";
        final Charset charset = Charsets.UTF_8;
        
        String xml;
        try {
            xml = getCreateListXml(charset, listName, listTemplate);
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Error creating list XML", ex);
            result.setException(ex);
            return result;
        }

        byte[] payload = EncodingUtils.getBytes(xml, charset.name());
        
        Map<String, String> headers = new HashMap<String, String>();
        headers.put("Content-Type", "application/atom+xml; charset=" + charset.name());
        
        ListenableFuture<JSONObject> request = executeRequestJson(createListUrl, "POST", headers, payload);
        
        Futures.addCallback(request, new FutureCallback<JSONObject>() {

            @Override
            public void onFailure(Throwable e) {
                result.setException(e);
            }

            @Override
            public void onSuccess(JSONObject json) {
                SPList list = new SPList();
                list.loadFromJson(json);
                result.set(list);
            }
        });
        
        return result;
    }
    
    private static class XmlNs {
        public static final String Atom = "http://www.w3.org/2005/Atom";
        public static final String DataServices = "http://schemas.microsoft.com/ado/2007/08/dataservices";
        public static final String DataServicesMetadata = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
    }

    private String getCreateListXml(Charset charset, String listName, String listTemplate) throws Exception {

        try {
            XmlSerializer xml = Xml.newSerializer();
            StringWriter sw = new StringWriter();
            xml.setOutput(sw);
            //document
            xml.startDocument(charset.name(), false);
            //namespaces
            xml.setPrefix("m", XmlNs.DataServicesMetadata);
            xml.setPrefix("d", XmlNs.DataServices);
            xml.setPrefix("a", XmlNs.Atom);
            {
                xml
                .startTag(XmlNs.Atom, "entry")
                    .startTag(XmlNs.Atom, "category")
                        .attribute(null, "term", "SP.List")
                        .attribute(null, "scheme", "http://schemas.microsoft.com/ado/2007/08/dataservices/scheme")
                    .endTag(XmlNs.Atom, "category")
                    .startTag(XmlNs.Atom, "content")
                        .attribute(null, "type", "application/xml")
                        .startTag(XmlNs.DataServicesMetadata, "properties")
                            .startTag(XmlNs.DataServices, "Title")
                                .text(listName)
                            .endTag(XmlNs.DataServices, "Title")
                            .startTag(XmlNs.DataServices, "BaseTemplate")
                                .text(listTemplate)
                            .endTag(XmlNs.DataServices, "BaseTemplate")
                        .endTag(XmlNs.DataServicesMetadata, "properties")
                    .endTag(XmlNs.Atom, "content")
                .endTag(XmlNs.Atom, "entry");
            }
            xml.endDocument();
            
            return sw.toString();
        }
        catch (Exception e) {
            throw new Exception("Error generating XML payload", e);
        }
    }
}
