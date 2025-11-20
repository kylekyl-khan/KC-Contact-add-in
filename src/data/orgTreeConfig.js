// 組織樹白名單設定：僅顯示指定的 Entra 群組
// 這裡維持 CommonJS 匯出，方便前端 bundle 以及 Node 腳本共用
const orgTreeConfig = [
  {
    id: "kangqiao",
    name: "康橋學校",
    children: [
      {
        id: "kcqs",
        name: "青山校區",
        children: [
          {
            id: "KCQS10",
            name: "青山校長室",
            children: [
              {
                id: "KCQS100101",
                name: "青山校長室人會組",
                // 對應 exportGroups_2025-11-19.csv 的 id 欄位
                groupId: "899b4fc4-1659-4666-8870-a1379977946a",
              },
            ],
          },
          {
            id: "KCQS1010",
            name: "青山教務處",
            children: [
              {
                id: "KCQS101001",
                name: "教學組",
                groupId: "6d32df83-c5ad-4adf-bb36-c34f4c1193d2",
              },
              {
                id: "KCQS101002",
                name: "課研組",
                groupId: "d3edc110-9105-4e45-b7d2-08a75f98c821",
              },
              {
                id: "KCQS101003",
                name: "課務組",
                groupId: "1578e794-48bd-423e-87f4-20ad6cba7eb5",
              },
              {
                id: "KCQS101004",
                name: "招生組",
                groupId: "2fcc5a35-542c-4ea0-ae76-82f708417b8e",
              },
            ],
          },
        ],
      },
      {
        id: "kccs",
        name: "常熟校區",
        children: [
          {
            id: "KCCS701307",
            name: "常熟招生處",
            children: [
              {
                id: "KCCS70130701",
                name: "招生一組",
                groupId: "be034064-8fd3-4ea0-b144-617a7315089f",
              },
              {
                id: "KCCS70130702",
                name: "招生二組",
                groupId: "9b8b1fe3-cb20-4187-a50b-0f4dd120e53a",
              },
            ],
          },
        ],
      },
    ],
  },
];

module.exports = {
  orgTreeConfig,
};
