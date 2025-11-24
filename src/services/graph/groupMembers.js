const { getGraphToken } = require("../auth/msalClient");

/**
 * 透過 Microsoft Graph 取得指定群組的成員
 * TODO: 若成員過多會收到 @odata.nextLink，之後需補上分頁處理。
 * @param {string} groupId
 * @returns {Promise<Array<{id:string,name:string,email:string,title:string}>>}
 */
async function getGroupMembers(groupId) {
  if (!groupId) return [];

  const token = await getGraphToken();
  const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,mail,jobTitle`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${errorText}`);
  }

  const data = await response.json();
  if (!data || !Array.isArray(data.value)) {
    throw new Error("Unexpected Graph response structure when reading members");
  }

  return data.value.map(member => ({
    id: member.id,
    name: member.displayName || "",
    email: member.mail || "",
    title: member.jobTitle || "",
  }));
}

module.exports = {
  getGroupMembers,
};
