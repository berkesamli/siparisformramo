// ========== ÇERÇEVE GÖRSELLERİ (SKU -> görsel ve kalınlık) ==========
// Önizlemede seri+renk kodu bu listedeki bir SKU ile eşleşirse çerçevenin
// gerçek profil fotoğrafı border olarak kullanılır; eşleşmeyenler koyu
// düz border ile gösterilir (görseller zamanla eklenecek).
//
// - clipRatio: Görselin çerçeve kenarından ne kadar kesileceği (0-1)
//   0.65 = varsayılan; küçük (0.4-0.5) ince, büyük (0.7-0.85) kalın çerçeveler
// - slice: border-image-slice yüzdesi (varsayılan FRAME_SLICE)
// - bareFrame: cam/paspartusuz çerçeveler (kanvas vb.)

export interface FrameImage {
  url: string;
  clipRatio: number;
  slice?: string;
  bareFrame?: boolean;
}

const u = (id: string) =>
  `https://cdn.myikas.com/images/04a76b35-2c55-499a-b485-0058f5ce13ce/${id}/image_1080.webp`;

export const FRAME_SLICE = "21%";
export const FRAME_BORDER_SCALE = 1.3;
export const DEFAULT_CLIP_RATIO = 0.65;

export const FRAME_IMAGES: Record<string, FrameImage> = {
  "GD154-4313-BA": { url: u("644a2fd0-366d-4a4c-9f9b-e01625c1afac"), clipRatio: 0.65 },
  "GD154-3427-BA": { url: u("80825c69-5301-42d2-a5ea-de83ba70ef24"), clipRatio: 0.65 },
  "GB139-1211T": { url: u("788a4a05-fc4c-4bd8-9b10-225338306230"), clipRatio: 0.73 },
  "GB139-2967T": { url: u("4ead5839-1fb1-41e0-9328-38a26a135da5"), clipRatio: 0.73 },
  "GB139-1775T": { url: u("1f42619f-0657-44c8-802f-c16101539b74"), clipRatio: 0.73 },
  "GG128-3110-P": { url: u("476aee5a-3b63-4802-8470-3c7bf46d506b"), clipRatio: 0.65 },
  "GB139-1921T": { url: u("dc847a75-079f-4ebb-9513-831f4330aa4a"), clipRatio: 0.73 },
  "GB185-1842-B": { url: u("2cffd38b-4038-45da-a2d4-5438da6447c3"), clipRatio: 0.73 },
  "GB185-3920-B": { url: u("da7ddf47-8d55-46e1-ae73-a90c170dcce5"), clipRatio: 0.73 },
  "GB185-4110-P": { url: u("3c5dbed4-62ec-46f2-a5d9-4190ae1f2f02"), clipRatio: 0.73 },
  "GB185-011-B": { url: u("a7f9f41c-5c04-43d7-8bfa-2ce477f1d24e"), clipRatio: 0.73 },
  "GG128-4346-BY": { url: u("8bc73df6-66fd-4566-971b-b0c1ce9fe075"), clipRatio: 0.65 },
  "GG128-3834-BY": { url: u("b856f27b-04a1-4f13-93b3-c1f74dca0604"), clipRatio: 0.65 },
  "GF141-4145-P": { url: u("22d4633b-738d-4ddb-b508-29145f13c8b9"), clipRatio: 0.6 },
  "GF141-3507-P": { url: u("ad847b33-85e0-4ae4-984f-60e9a1dd16ce"), clipRatio: 0.6 },
  "GC242-3925": { url: u("1152e7ed-96b3-48a3-a530-2b4012052ef0"), clipRatio: 0.7 },
  "GC242-3922": { url: u("09664d49-4ac4-452a-8207-ba7fe6496a2e"), clipRatio: 0.7 },
  "GE169-3115P": { url: u("aca41223-e4d7-4da7-9cf0-3b9fa1f34034"), clipRatio: 0.65 },
  "GE169-3110P": { url: u("fd23cbbc-116d-4798-9aab-04ba5c1a58dd"), clipRatio: 0.65 },
  "GE169-4132P": { url: u("b5dc2327-64b4-4c7f-ba8e-0065eff5e071"), clipRatio: 0.65 },
  "GE169-4133P": { url: u("32244866-0683-485a-8757-722c4a1c6be9"), clipRatio: 0.65 },
  "GE146-3289B": { url: u("bf5c951e-eb4c-4126-80f7-8d43892b48ec"), clipRatio: 0.65 },
  "GE146-3288B": { url: u("0c6e33b7-1e99-4d09-afab-93ef701ce87e"), clipRatio: 0.65 },
  "GC176-3289": { url: u("7c724926-5f6b-433b-907d-35400bcfd709"), clipRatio: 0.7 },
  "GC176-3288": { url: u("d002f399-e67f-4b2b-a0a5-f1c0fa1dbd94"), clipRatio: 0.7 },
  "GB066-4147": { url: u("f244843d-cd27-47b6-bbb8-0834a7867aa4"), clipRatio: 0.73 },
  "GB066-2211": { url: u("9e02e010-fb03-4a60-9a8c-8c48badc4f3e"), clipRatio: 0.73 },
  "GB066-2725": { url: u("18d6a13e-137f-4239-8662-45725e103462"), clipRatio: 0.73 },
  "GB066-3271": { url: u("5f054e67-7d0e-4d84-88d3-ee468a54e2e3"), clipRatio: 0.73 },
  "GB147-4147": { url: u("ac42dfc1-6c05-4e8d-8a08-2188d35c8c4f"), clipRatio: 0.7 },
  "GB147-4013B": { url: u("317c2baa-7de8-4309-ba52-73267fbf43d1"), clipRatio: 0.7 },
  "GB147-4017B": { url: u("2051a109-bac9-47b6-9282-917b9562b5d1"), clipRatio: 0.7 },
  "GB147-4016B": { url: u("ed2f56a5-afca-4167-a720-27cfde2ea46d"), clipRatio: 0.7 },
  "4313S-N78": { url: u("3b80cf75-6718-4135-96aa-c6bfe9737e0c"), clipRatio: 0.73 },
  "4313S-P78": { url: u("a600d9ee-23f0-4785-bce9-0df5fd2205b1"), clipRatio: 0.73 },
  "3248S-P78": { url: u("236dacfe-7648-4034-b555-38fdd71ffa97"), clipRatio: 0.73 },
  "3248S-X81G": { url: u("e4957475-3ddc-485e-a43b-3e65b8ec2961"), clipRatio: 0.73 },
  "3248S-S85Y": { url: u("a7d78a4f-5f51-4faf-bb89-07e12ac50dc9"), clipRatio: 0.73 },
  "3248S-N78": { url: u("d4921300-637c-4a39-b861-fbd251a1521f"), clipRatio: 0.73 },
  "2315S-N90": { url: u("b694d2a5-1e3d-4b9d-a0a4-ab3ff5d721a7"), clipRatio: 0.73 },
  "2315S-L90": { url: u("bbe412d7-4dfa-48c5-81f7-1a7f3dae797a"), clipRatio: 0.73 },
  "2315S-D1": { url: u("b22d14c6-75c5-4a54-adbf-7d8c3ecae52c"), clipRatio: 0.73 },
  "2315S-C75": { url: u("92af57c3-637c-4399-a17a-b453a57c443a"), clipRatio: 0.73 },
  // 2231 Serisi
  "2231-J67": { url: u("e36100c3-a9ac-44c1-9294-2f0fdaeb08b5"), clipRatio: 0.73 },
  "2231-R80": { url: u("de3b9743-fb15-41d5-923b-ed221d7c90b1"), clipRatio: 0.73 },
  "2231-F81": { url: u("edb1009f-951e-4c58-b056-a841d4e58ce0"), clipRatio: 0.73 },
  "2231-D1Y": { url: u("6cbf8651-83f0-4ffb-94db-35a328ca274f"), clipRatio: 0.73 },
  "FC062-489S": { url: u("f1e04c62-7b3b-4f3e-9cd3-a47583d943b7"), clipRatio: 0.73 },
  "FC062-486S": { url: u("4c5a13c1-e463-42cc-a602-001936ab7230"), clipRatio: 0.73 },
  "FC062-919S": { url: u("08745453-c412-40c5-bde9-43b9d8f68040"), clipRatio: 0.73 },
  "FC062-918S": { url: u("a6a4df61-8edc-4286-8f23-8de9109d3f90"), clipRatio: 0.73 },
  "FC062-917S": { url: u("dd4cbd31-d70e-42fe-bd1a-4798ac2c2283"), clipRatio: 0.73 },
  "FB056-025": { url: u("e4f114a3-b2d0-437f-9129-8aa255752860"), clipRatio: 0.73 },
  "FB056-014": { url: u("1a0d3c54-4479-4e81-a459-25ec45b6d68d"), clipRatio: 0.73 },
  "GG044-916BX": { url: u("66d696a5-5939-4d53-8ee6-a0386957b373"), clipRatio: 0.6 },
  "GG044-2013BX": { url: u("133f1a33-4672-46d7-b9ff-35250967165c"), clipRatio: 0.6 },
  "GG044-1863BX": { url: u("9a339c1e-2ff9-4cbd-a87d-9d58247cf546"), clipRatio: 0.6 },
  // GC116 Serisi
  "GC116-2214C": { url: u("9218d1b6-fbf8-44f2-b0b1-fdffa549f8e6"), clipRatio: 0.7 },
  "GC116-2213C": { url: u("238da11a-a6bb-46e1-8cd2-cee1cd91071f"), clipRatio: 0.7 },
  // GC148 Serisi
  "GC148-2701A": { url: u("6a189cd7-90f4-4c3a-9bac-3fc7c9bbba0f"), clipRatio: 0.7 },
  "GC148-2211A": { url: u("a8b4ae41-cd65-4308-b269-f55fbb0bdc18"), clipRatio: 0.7 },
  "GC148-2145A": { url: u("6c899f87-fff6-418a-8c9b-63d1088542c2"), clipRatio: 0.7 },
  "GC148-1565A": { url: u("872355cf-43ba-47f6-8045-0b47bcc9c9b9"), clipRatio: 0.7 },
  // GD004 Serisi
  "GD004-864BX": { url: u("964f756e-6187-4d94-ae02-96844a870d3d"), clipRatio: 0.68 },
  "GD004-033": { url: u("ecfb8b78-b3c7-4790-b18e-3fcc0afe7bb1"), clipRatio: 0.68 },
  "GD004-031": { url: u("18852bbc-d33f-42af-a695-39d340fc9536"), clipRatio: 0.68 },
  "GD004-017": { url: u("12941e89-1115-4f6a-b6af-57c8f4de3929"), clipRatio: 0.68 },
  "GD004-011BX": { url: u("624f3953-e289-4708-9f20-7ad549dadc52"), clipRatio: 0.68 },
  // GD138 Serisi
  "GD138-2997B": { url: u("641fb028-628b-4ade-8c04-12598eb0af00"), clipRatio: 0.68 },
  "GD138-2845B": { url: u("551ef47a-87ff-42e5-985a-a263f70ad57e"), clipRatio: 0.68 },
  "GD138-2842B": { url: u("cbe934ef-5ae8-446b-b4aa-91ec05055f25"), clipRatio: 0.68 },
  "GD138-2841B": { url: u("47802748-8102-4301-bc01-755531ea7a79"), clipRatio: 0.68 },
  // GE015 Serisi
  "GE015-481BX": { url: u("7ee0b8cf-43aa-4ecd-a295-3ffda92bb233"), clipRatio: 0.65 },
  "GE015-480BX": { url: u("2c49ae3a-cc6f-415b-8de6-c191f47cd134"), clipRatio: 0.65 },
  "GE015-2013BX": { url: u("8470a18d-35d4-48bd-b3aa-4f3028b94695"), clipRatio: 0.65 },
  "GE015-1863BX": { url: u("0a0c64fc-51f4-4446-9b31-ebecdb30b16e"), clipRatio: 0.65 },
  // GE070 Serisi
  "GE070-81467B": { url: u("f75ac52d-956f-4da7-819b-fbaa3327e7f5"), clipRatio: 0.65 },
  "GE070-633B": { url: u("567eee4c-a4e4-4524-8a50-3d6eeb30457f"), clipRatio: 0.65 },
  "GE070-1210": { url: u("5e164228-310b-4c91-ab7b-6c6e3e437fc5"), clipRatio: 0.65 },
  "GE070-1039B": { url: u("e4162a45-0e0d-45e5-9c48-b48184db62cf"), clipRatio: 0.65 },
  "GE070-1038B": { url: u("94c2d4de-d86e-47a4-80ad-86edf6b4f30e"), clipRatio: 0.65 },
  // GE105 Serisi
  "GE105-2083T": { url: u("c4342584-eb20-4597-8f58-8b20b5301fab"), clipRatio: 0.65 },
  "GE105-2056T": { url: u("101c93b2-f341-4a4e-95d3-7e0ce90474f1"), clipRatio: 0.65 },
  // GF082 Serisi
  "GF082-2092T": { url: u("0cedd9e4-bbaf-4e90-95dd-a79d4094ddc2"), clipRatio: 0.6 },
  "GF082-2013T": { url: u("2d6768e2-a93a-462b-8671-c0a6846c49f0"), clipRatio: 0.6 },
  "GF082-1863T": { url: u("7efbeab3-8a19-4bd6-848c-6d7a444118a1"), clipRatio: 0.6 },
  // GF105 Serisi
  "GF105-2701A": { url: u("224117c5-df90-4237-b9ab-647e7cd7e027"), clipRatio: 0.6 },
  "GF105-2211A": { url: u("a6d451d4-7f2c-44fe-a18e-f7378b9397e9"), clipRatio: 0.6 },
  "GF105-2145A": { url: u("f6d8afe1-cbd1-4720-b5cf-2f55e1483c58"), clipRatio: 0.6 },
  "GF105-1565A": { url: u("ae2643e7-73d1-49a6-9381-68fad1ba96be"), clipRatio: 0.6 },
  // FB053 Serisi
  "FB053-964": { url: u("e61ec6da-2934-483d-99b1-1b29cb30c215"), clipRatio: 0.72 },
  "FB053-962S": { url: u("3fca6e2a-6e88-46ba-a927-53be991c78c8"), clipRatio: 0.72 },
  "FB053-963S": { url: u("44c45810-efac-4c1b-894b-5225f9be4e18"), clipRatio: 0.72 },
  "FB053-959": { url: u("ca4834c7-176e-4d8b-bd2a-12505d90fb68"), clipRatio: 0.72 },
  // FB037 Serisi
  "FB037-576": { url: u("674f6f56-3bed-4003-b3d0-2f929fc19d02"), clipRatio: 0.77 },
  "FB037-576S": { url: u("fe87d63a-b315-4bbd-a00d-8a12cda758bf"), clipRatio: 0.77 },
  "FB037-760S": { url: u("97a8f0a3-2b69-4768-81bb-5230c5fe9d20"), clipRatio: 0.77 },
  "FB037-014S": { url: u("7d2dd0b8-6db4-4a57-876d-bb3562ecc926"), clipRatio: 0.77 },
  "FB037-025S": { url: u("8e46e225-78e2-4787-9ca5-2f23944b820a"), clipRatio: 0.77 },
  // FB022 Serisi
  "FB022-595S": { url: u("06c7721d-3a69-419f-ab8a-27998e2f05e3"), clipRatio: 0.73 },
  "FB022-628S": { url: u("f0fed62b-905b-41a7-ac57-7fc425c5b098"), clipRatio: 0.73 },
  "FB022-032S": { url: u("343e654f-f897-4a8b-b3e7-c0b1f79cb722"), clipRatio: 0.73 },
  // FB014 Serisi
  "FB014-330": { url: u("a3edbd88-1f5e-4a53-95ab-0ea011a49b85"), clipRatio: 0.73 },
  "FB014-256": { url: u("d18c1569-0df6-42ca-8ff7-6f4de978602f"), clipRatio: 0.73 },
  "FB014-598": { url: u("94fef9dc-6b0f-4d7a-b226-62dd5c161455"), clipRatio: 0.73 },
  "FB014-014": { url: u("086d28a4-9215-4439-9b19-acb68901ad63"), clipRatio: 0.73 },
  "FB014-025": { url: u("281f2f5a-7982-454a-a60d-2b9c1d6c1fe7"), clipRatio: 0.73 },
  // FB002 Serisi
  "FB002-760S": { url: u("97a8f0a3-2b69-4768-81bb-5230c5fe9d20"), clipRatio: 0.77 },
  "FB002-014S": { url: u("7d2dd0b8-6db4-4a57-876d-bb3562ecc926"), clipRatio: 0.77 },
  "FB002-025S": { url: u("8e46e225-78e2-4787-9ca5-2f23944b820a"), clipRatio: 0.77 },
  // GG090 Serisi
  "GG090-2851T": { url: u("119b82cf-2d5f-4f9a-a406-57ebf50178d6"), clipRatio: 0.58, slice: "25%" },
  "GG090-2850T": { url: u("4ea1fc01-6d65-40fd-a3d0-64ca9ca6d975"), clipRatio: 0.58, slice: "25%" },
  // GG048 Serisi
  "GG048-1039B": { url: u("cd55f13d-d18e-400a-827b-7a18bc1726d8"), clipRatio: 0.6 },
  "GG048-1038B": { url: u("06f65bcb-6a79-473b-aac2-bae452cf96d6"), clipRatio: 0.6 },
  "GG048-1210": { url: u("f507dab5-c352-4541-a84c-ef2e1eef53b9"), clipRatio: 0.6 },
  // Yeni çerçeve eklerken: "SKU-KODU": { url: u("<görsel-id>"), clipRatio: 0.65 },

  // ---- Cam/paspartu olmayan çerçeveler (bareFrame) ----
  "GC153-011A": { url: u("94d0e324-ce55-48d6-8623-93190b385b5c"), clipRatio: 0.65, bareFrame: true },
  "GC153-2927A": { url: u("74e99f6b-e8ea-4504-8154-864000a55c93"), clipRatio: 0.65, bareFrame: true },
  "GC153-2926A": { url: u("ace52482-5a8e-4e99-b124-96959c0ecbc6"), clipRatio: 0.65, bareFrame: true },
  "GC153-2116A": { url: u("4ee0c527-836a-4527-97b3-ea2f5b4a0577"), clipRatio: 0.65, bareFrame: true },
  // GC056 Serisi
  "GC056-843B": { url: u("45d55d19-bd55-4144-9de1-7f8e6156bb7f"), clipRatio: 0.7 },
  "GC056-1110B": { url: u("1e90687c-24fd-4456-b155-b9274259cb4c"), clipRatio: 0.7 },
  // GC017 Serisi
  "GC017-707A": { url: u("67920bf2-cf56-4bb6-b767-a5a70e1b791d"), clipRatio: 0.7 },
  "GC017-2935A": { url: u("b5edd2ef-8b8a-4e9f-988c-8fc7ae462ebc"), clipRatio: 0.7 },
  "GC017-2934A": { url: u("4396660b-b84f-494c-a966-8a8f56ff3404"), clipRatio: 0.7 },
  // GB119 Serisi
  "GB119-2498": { url: u("ba09a36d-74d1-40f9-8632-18b306443db3"), clipRatio: 0.7 },
  "GB119-2496T": { url: u("06d89b7c-33de-4dba-bb4b-466b6e147f89"), clipRatio: 0.7 },
  "GB119-2476": { url: u("fcbaca0c-c177-4117-9e5f-1e86d8ea7f9b"), clipRatio: 0.7 },
  "GB119-2475T": { url: u("730e7642-0c25-4b75-97d5-c3d6c13e18b0"), clipRatio: 0.7 },
  // GB108 Serisi
  "GB108-1489BC": { url: u("a03799ce-d27d-4621-9daf-b42a626013e9"), clipRatio: 0.72 },
  "GB108-2211B": { url: u("bdcde2e3-60e8-4f51-92f8-13e5bda90e41"), clipRatio: 0.72 },
  "GB108-033": { url: u("2b0e2835-5c84-45f1-865f-8f05154adb6e"), clipRatio: 0.72 },
  "GB108-017": { url: u("79da7683-942e-45a3-a008-bcf3a7e863e4"), clipRatio: 0.72 },
  "GB108-011BX": { url: u("f9f8561b-7f3b-4c99-94a0-872d57ef1b43"), clipRatio: 0.72 },
  // GB085 Serisi
  "GB085-1633B": { url: u("605346f8-0085-47f0-a501-dbd79d06e4df"), clipRatio: 0.73 },
  "GB085-1490B": { url: u("c4276ba2-c5ab-41df-b8cd-c6d338fd5ea5"), clipRatio: 0.73 },
  // GB022 Serisi
  "GB022-864BX": { url: u("a4145e13-61f1-41ff-8c6e-a169a95223ec"), clipRatio: 0.75 },
  "GB022-501NY": { url: u("c7dc3b53-89a3-47e4-b0f4-9145afd8e089"), clipRatio: 0.75 },
  "GB022-500NY": { url: u("c020563a-f265-4f98-ace1-7c55e3f18701"), clipRatio: 0.75 },
  "GB022-298BX": { url: u("2975f1b4-3a74-4616-bfde-1855b404596a"), clipRatio: 0.75 },
  "GB022-1542BX": { url: u("cef68fd6-8272-4310-a3b8-c8b937537e51"), clipRatio: 0.75 },
  "GB022-030BX": { url: u("993367d1-74a7-4c6a-aa56-b730c5aef377"), clipRatio: 0.75 },
  "GB022-011BX": { url: u("0a240855-6d2b-4b05-960f-f343017d4984"), clipRatio: 0.75 },
  // AH Serisi
  "AH-WT": { url: u("f53d6e71-726d-42dd-9eb3-c08687c88d6a"), clipRatio: 0.75 },
  "AH-BK": { url: u("5c6c1a2a-ef00-4e34-a6a2-9fffce7c350b"), clipRatio: 0.75 },
  "AH-07": { url: u("fe3b7d1c-e56e-40b9-8296-4bb7451abaa3"), clipRatio: 0.75 },
  // FD070 Serisi (Canvas)
  "FD070-355": { url: u("90011e9a-f020-4337-b50a-5521ac76a558"), clipRatio: 0.65, bareFrame: true },
  "FD070-1570": { url: u("5d803bbc-e093-42b3-a70c-56b208c269fd"), clipRatio: 0.65, bareFrame: true },
  "FD070-025": { url: u("a78cf9c8-6a9a-4ed2-9e2d-aa3ee83172d3"), clipRatio: 0.65, bareFrame: true },
};

// SKU eşleştirme: büyük/küçük harf ve boşluk duyarsız
const norm = (s: string) => s.toUpperCase().replace(/\s+/g, "").trim();

const INDEX: Record<string, FrameImage> = {};
for (const [sku, img] of Object.entries(FRAME_IMAGES)) {
  INDEX[norm(sku)] = img;
}

export function findFrameImage(fullCode: string): FrameImage | null {
  if (!fullCode) return null;
  return INDEX[norm(fullCode)] || null;
}
