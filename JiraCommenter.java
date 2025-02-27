
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;

public class JiraService {

    private static final String JIRA_BASE_URL = "https://your-jira-instance.atlassian.net"; // Change to your Jira URL
    private static final String JIRA_USERNAME = "your-email@example.com"; // Change to your Jira username
    private static final String JIRA_API_TOKEN = "your-api-token"; // Change to your Jira API token
    private static final String ISSUE_KEY = "PROJ-123"; // Change to your issue key

    public static void main(String[] args) {
        try {
            addCommentToJiraIssue(ISSUE_KEY, "This is a test comment from Java.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void addCommentToJiraIssue(String issueKey, String comment) throws Exception {
        String url = JIRA_BASE_URL + "/rest/api/3/issue/" + issueKey + "/comment";

        CloseableHttpClient client = HttpClients.createDefault();
        HttpPost httpPost = new HttpPost(url);

        // Authentication
        String auth = JIRA_USERNAME + ":" + JIRA_API_TOKEN;
        String encodedAuth = Base64.getEncoder().encodeToString(auth.getBytes(StandardCharsets.UTF_8));
        httpPost.setHeader("Authorization", "Basic " + encodedAuth);
        httpPost.setHeader("Content-Type", "application/json");

        // Create JSON payload
        Map<String, Object> body = new HashMap<>();
        body.put("body", comment);

        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPayload = objectMapper.writeValueAsString(body);

        httpPost.setEntity(new StringEntity(jsonPayload, StandardCharsets.UTF_8));

        // Execute request
        CloseableHttpResponse response = client.execute(httpPost);
        String responseBody = EntityUtils.toString(response.getEntity());

        System.out.println("Response Code: " + response.getStatusLine().getStatusCode());
        System.out.println("Response Body: " + responseBody);

        client.close();
    }
}
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;

public class JiraCommenter {

    private static final String JIRA_BASE_URL = "https://your-jira-instance.atlassian.net"; // Change to your Jira URL
    private static final String JIRA_USERNAME = "your-email@example.com"; // Change to your Jira username
    private static final String JIRA_API_TOKEN = "your-api-token"; // Change to your Jira API token
    private static final String ISSUE_KEY = "PROJ-123"; // Change to your issue key

    public static void main(String[] args) {
        try {
            addCommentToJiraIssue(ISSUE_KEY, "This is a test comment from Java.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void addCommentToJiraIssue(String issueKey, String comment) throws Exception {
        String url = JIRA_BASE_URL + "/rest/api/3/issue/" + issueKey + "/comment";

        CloseableHttpClient client = HttpClients.createDefault();
        HttpPost httpPost = new HttpPost(url);

        // Authentication
        String auth = JIRA_USERNAME + ":" + JIRA_API_TOKEN;
        String encodedAuth = Base64.getEncoder().encodeToString(auth.getBytes(StandardCharsets.UTF_8));
        httpPost.setHeader("Authorization", "Basic " + encodedAuth);
        httpPost.setHeader("Content-Type", "application/json");

        // Create JSON payload
        Map<String, Object> body = new HashMap<>();
        body.put("body", comment);

        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPayload = objectMapper.writeValueAsString(body);

        httpPost.setEntity(new StringEntity(jsonPayload, StandardCharsets.UTF_8));

        // Execute request
        CloseableHttpResponse response = client.execute(httpPost);
        String responseBody = EntityUtils.toString(response.getEntity());

        System.out.println("Response Code: " + response.getStatusLine().getStatusCode());
        System.out.println("Response Body: " + responseBody);

        client.close();
    }
}

