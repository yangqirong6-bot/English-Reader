'use client';

import { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import { speakSentence, stopPlayback } from '@/lib/azure-tts';
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from 'docx';
import Draggable from 'react-draggable';
import type { DictResult } from '@/lib/dict-types';

type DictApiResponse = DictResult | { html: string };

// ─── Types ───────────────────────────────────────────
interface WordToken {
  display: string;
  clean: string;
  id: string;
}

interface SentenceData {
  words: WordToken[];
  id: string;
}

interface DictEntry {
  phonetic: string;
  en: string;
  zh: string;
}

interface TooltipState {
  word: WordToken;
  entry: DictEntry | null;
  htmlDefinition: string | null;
  top: number;
  left: number;
}

// ─── Mock dictionary ─────────────────────────────────
const MOCK_DICT: Record<string, DictEntry> = {
  // Common prepositions & articles
  the: { phonetic: '/ðə/', en: 'used to refer to a specific noun', zh: '（定冠词）这，那' },
  a: { phonetic: '/ə/', en: 'used before a singular noun', zh: '一个' },
  an: { phonetic: '/ən/', en: 'used before a vowel sound', zh: '一个' },
  in: { phonetic: '/ɪn/', en: 'expressing location or time', zh: '在…里面' },
  on: { phonetic: '/ɒn/', en: 'physically in contact with a surface', zh: '在…上面' },
  at: { phonetic: '/æt/', en: 'expressing location or time', zh: '在（某处或某时）' },
  to: { phonetic: '/tuː/', en: 'expressing direction or purpose', zh: '向，到' },
  for: { phonetic: '/fɔːr/', en: 'indicating purpose or recipient', zh: '为了，给' },
  with: { phonetic: '/wɪð/', en: 'accompanied by', zh: '和…一起' },
  from: { phonetic: '/frʌm/', en: 'indicating origin or source', zh: '从，来自' },
  by: { phonetic: '/baɪ/', en: 'indicating the agent or means', zh: '通过，由' },
  as: { phonetic: '/æz/', en: 'used to indicate function or comparison', zh: '作为，如同' },
  of: { phonetic: '/ɒv/', en: 'expressing relationship or belonging', zh: '…的' },

  // Common verbs
  is: { phonetic: '/ɪz/', en: 'third person singular of "be"', zh: '是' },
  are: { phonetic: '/ɑːr/', en: 'plural present of "be"', zh: '是' },
  was: { phonetic: '/wɒz/', en: 'first/third person past of "be"', zh: '（过去式）是' },
  were: { phonetic: '/wɜːr/', en: 'plural past of "be"', zh: '（过去式）是' },
  have: { phonetic: '/hæv/', en: 'to possess or own', zh: '有' },
  has: { phonetic: '/hæz/', en: 'third person singular of "have"', zh: '有' },
  had: { phonetic: '/hæd/', en: 'past tense of "have"', zh: '曾经有' },
  do: { phonetic: '/duː/', en: 'to perform an action', zh: '做' },
  does: { phonetic: '/dʌz/', en: 'third person singular of "do"', zh: '做' },
  did: { phonetic: '/dɪd/', en: 'past tense of "do"', zh: '做了' },
  say: { phonetic: '/seɪ/', en: 'to speak or utter', zh: '说' },
  says: { phonetic: '/sez/', en: 'third person singular of "say"', zh: '说' },
  said: { phonetic: '/sed/', en: 'past tense of "say"', zh: '说过' },
  go: { phonetic: '/ɡəʊ/', en: 'to move or travel', zh: '去' },
  get: { phonetic: '/ɡet/', en: 'to obtain or receive', zh: '得到' },
  make: { phonetic: '/meɪk/', en: 'to create or produce', zh: '制作，使' },
  made: { phonetic: '/meɪd/', en: 'past tense of "make"', zh: '制作了' },
  take: { phonetic: '/teɪk/', en: 'to lay hold of', zh: '拿，取' },
  see: { phonetic: '/siː/', en: 'to perceive with the eyes', zh: '看见' },
  come: { phonetic: '/kʌm/', en: 'to approach or arrive', zh: '来' },
  know: { phonetic: '/nəʊ/', en: 'to be aware of through observation', zh: '知道' },
  think: { phonetic: '/θɪŋk/', en: 'to have a belief or opinion', zh: '认为，想' },
  want: { phonetic: '/wɒnt/', en: 'to desire or wish for', zh: '想要' },
  give: { phonetic: '/ɡɪv/', en: 'to freely transfer possession', zh: '给' },
  use: { phonetic: '/juːz/', en: 'to employ for a purpose', zh: '使用' },
  find: { phonetic: '/faɪnd/', en: 'to discover or locate', zh: '找到' },
  tell: { phonetic: '/tel/', en: 'to communicate information', zh: '告诉' },
  ask: { phonetic: '/ɑːsk/', en: 'to pose a question', zh: '问' },
  work: { phonetic: '/wɜːk/', en: 'to perform a task or job', zh: '工作' },
  seem: { phonetic: '/siːm/', en: 'to give the impression of', zh: '似乎' },
  feel: { phonetic: '/fiːl/', en: 'to experience an emotion', zh: '感觉' },
  try: { phonetic: '/traɪ/', en: 'to attempt or test', zh: '尝试' },
  leave: { phonetic: '/liːv/', en: 'to go away from', zh: '离开' },
  call: { phonetic: '/kɔːl/', en: 'to give a name or contact by phone', zh: '打电话，称呼' },
  need: { phonetic: '/niːd/', en: 'to require something essential', zh: '需要' },
  mean: { phonetic: '/miːn/', en: 'to intend to convey', zh: '意思是' },
  keep: { phonetic: '/kiːp/', en: 'to continue to have', zh: '保持' },
  let: { phonetic: '/let/', en: 'to allow or permit', zh: '让' },
  begin: { phonetic: '/bɪˈɡɪn/', en: 'to start or commence', zh: '开始' },
  become: { phonetic: '/bɪˈkʌm/', en: 'to turn into', zh: '变成' },
  show: { phonetic: '/ʃəʊ/', en: 'to display or demonstrate', zh: '展示' },
  hear: { phonetic: '/hɪər/', en: 'to perceive with the ear', zh: '听到' },
  play: { phonetic: '/pleɪ/', en: 'to engage in an activity for fun', zh: '玩，播放' },
  run: { phonetic: '/rʌn/', en: 'to move at a fast pace', zh: '跑' },
  move: { phonetic: '/muːv/', en: 'to change position', zh: '移动' },
  live: { phonetic: '/lɪv/', en: 'to be alive or reside', zh: '生活，居住' },
  believe: { phonetic: '/bɪˈliːv/', en: 'to accept as true', zh: '相信' },
  bring: { phonetic: '/brɪŋ/', en: 'to carry toward the speaker', zh: '带来' },
  happen: { phonetic: '/ˈhæpən/', en: 'to occur or take place', zh: '发生' },
  write: { phonetic: '/raɪt/', en: 'to compose text', zh: '写' },
  provide: { phonetic: '/prəˈvaɪd/', en: 'to supply or make available', zh: '提供' },
  sit: { phonetic: '/sɪt/', en: 'to be seated', zh: '坐' },
  stand: { phonetic: '/stænd/', en: 'to be upright on the feet', zh: '站' },
  lose: { phonetic: '/luːz/', en: 'to be deprived of', zh: '失去' },
  pay: { phonetic: '/peɪ/', en: 'to give money for goods', zh: '支付' },
  meet: { phonetic: '/miːt/', en: 'to encounter or come together', zh: '遇见' },
  include: { phonetic: '/ɪnˈkluːd/', en: 'to contain as part of', zh: '包括' },
  continue: { phonetic: '/kənˈtɪnjuː/', en: 'to keep doing without stopping', zh: '继续' },
  set: { phonetic: '/set/', en: 'to put in a specific place', zh: '设置，放置' },
  learn: { phonetic: '/lɜːn/', en: 'to gain knowledge or skill', zh: '学习' },
  change: { phonetic: '/tʃeɪndʒ/', en: 'to make different', zh: '改变' },
  lead: { phonetic: '/liːd/', en: 'to guide or direct', zh: '带领' },
  understand: { phonetic: '/ˌʌndəˈstænd/', en: 'to comprehend', zh: '理解' },
  watch: { phonetic: '/wɒtʃ/', en: 'to observe attentively', zh: '观看' },
  follow: { phonetic: '/ˈfɒləʊ/', en: 'to come after or pursue', zh: '跟随' },
  stop: { phonetic: '/stɒp/', en: 'to cease moving or operating', zh: '停止' },
  create: { phonetic: '/kriˈeɪt/', en: 'to bring into existence', zh: '创造' },
  speak: { phonetic: '/spiːk/', en: 'to talk or utter words', zh: '说话' },
  read: { phonetic: '/riːd/', en: 'to interpret written text', zh: '阅读' },
  allow: { phonetic: '/əˈlaʊ/', en: 'to permit or give permission', zh: '允许' },
  add: { phonetic: '/æd/', en: 'to combine or join', zh: '添加' },
  spend: { phonetic: '/spend/', en: 'to use time or money', zh: '花费' },
  grow: { phonetic: '/ɡrəʊ/', en: 'to increase in size', zh: '成长' },
  open: { phonetic: '/ˈəʊpən/', en: 'to move to an accessible position', zh: '打开' },
  walk: { phonetic: '/wɔːk/', en: 'to move on foot', zh: '步行' },
  win: { phonetic: '/wɪn/', en: 'to achieve victory', zh: '赢得' },
  offer: { phonetic: '/ˈɒfər/', en: 'to present for acceptance', zh: '提供' },
  remember: { phonetic: '/rɪˈmembər/', en: 'to recall to mind', zh: '记住' },
  consider: { phonetic: '/kənˈsɪdər/', en: 'to think carefully about', zh: '考虑' },
  appear: { phonetic: '/əˈpɪər/', en: 'to come into sight', zh: '出现' },
  buy: { phonetic: '/baɪ/', en: 'to purchase with money', zh: '购买' },
  serve: { phonetic: '/sɜːv/', en: 'to perform duties for', zh: '服务' },
  die: { phonetic: '/daɪ/', en: 'to cease living', zh: '死亡' },
  send: { phonetic: '/send/', en: 'to cause to be delivered', zh: '发送' },
  build: { phonetic: '/bɪlt/', en: 'to construct', zh: '建造' },
  stay: { phonetic: '/steɪ/', en: 'to remain in the same place', zh: '停留' },
  fall: { phonetic: '/fɔːl/', en: 'to drop downward', zh: '落下' },
  cut: { phonetic: '/kʌt/', en: 'to divide with a sharp tool', zh: '切割' },
  reach: { phonetic: '/riːtʃ/', en: 'to stretch to touch', zh: '到达' },
  kill: { phonetic: '/kɪl/', en: 'to cause death', zh: '杀死' },
  remain: { phonetic: '/rɪˈmeɪn/', en: 'to continue to exist', zh: '保持，剩余' },
  suggest: { phonetic: '/səˈdʒest/', en: 'to put forward for consideration', zh: '建议' },
  raise: { phonetic: '/reɪz/', en: 'to lift upward', zh: '提高，举起' },
  expect: { phonetic: '/ɪkˈspekt/', en: 'to regard as likely to happen', zh: '期待' },

  // Nouns
  time: { phonetic: '/taɪm/', en: 'the indefinite continued progress of existence', zh: '时间' },
  year: { phonetic: '/jɪər/', en: 'the period of 365 days', zh: '年' },
  people: { phonetic: '/ˈpiːpl/', en: 'human beings in general', zh: '人们' },
  way: { phonetic: '/weɪ/', en: 'a method or manner of doing something', zh: '方式，路' },
  day: { phonetic: '/deɪ/', en: 'a 24-hour period', zh: '天' },
  thing: { phonetic: '/θɪŋ/', en: 'an object or entity', zh: '东西' },
  man: { phonetic: '/mæn/', en: 'an adult male human', zh: '男人' },
  woman: { phonetic: '/ˈwʊmən/', en: 'an adult female human', zh: '女人' },
  child: { phonetic: '/tʃaɪld/', en: 'a young human being', zh: '孩子' },
  world: { phonetic: '/wɜːld/', en: 'the earth and all its inhabitants', zh: '世界' },
  life: { phonetic: '/laɪf/', en: 'the condition of being alive', zh: '生活，生命' },
  hand: { phonetic: '/hænd/', en: 'the end part of the arm', zh: '手' },
  part: { phonetic: '/pɑːt/', en: 'a piece or segment', zh: '部分' },
  place: { phonetic: '/pleɪs/', en: 'a particular position or location', zh: '地方' },
  case: { phonetic: '/keɪs/', en: 'an instance or occurrence', zh: '情况，案例' },
  week: { phonetic: '/wiːk/', en: 'a period of seven days', zh: '周' },
  company: { phonetic: '/ˈkʌmpəni/', en: 'a business organization', zh: '公司' },
  system: { phonetic: '/ˈsɪstəm/', en: 'a set of connected parts', zh: '系统' },
  program: { phonetic: '/ˈprəʊɡræm/', en: 'a planned series of events', zh: '计划，程序' },
  question: { phonetic: '/ˈkwestʃən/', en: 'an inquiry or query', zh: '问题' },
  government: { phonetic: '/ˈɡʌvənmənt/', en: 'the governing body of a nation', zh: '政府' },
  number: { phonetic: '/ˈnʌmbər/', en: 'a count or quantity', zh: '数字' },
  night: { phonetic: '/naɪt/', en: 'the period of darkness', zh: '夜晚' },
  point: { phonetic: '/pɔɪnt/', en: 'a specific detail or location', zh: '点，要点' },
  home: { phonetic: '/həʊm/', en: 'the place where one lives', zh: '家' },
  water: { phonetic: '/ˈwɔːtər/', en: 'a transparent liquid essential for life', zh: '水' },
  room: { phonetic: '/ruːm/', en: 'an area of a building', zh: '房间' },
  mother: { phonetic: '/ˈmʌðər/', en: 'a female parent', zh: '母亲' },
  father: { phonetic: '/ˈfɑːðər/', en: 'a male parent', zh: '父亲' },
  family: { phonetic: '/ˈfæməli/', en: 'a group of related people', zh: '家庭' },
  school: { phonetic: '/skuːl/', en: 'an educational institution', zh: '学校' },
  state: { phonetic: '/steɪt/', en: 'a condition or territory', zh: '状态，州' },
  eye: { phonetic: '/aɪ/', en: 'the organ of sight', zh: '眼睛' },
  head: { phonetic: '/hed/', en: 'the upper part of the body', zh: '头' },
  group: { phonetic: '/ɡruːp/', en: 'a number of people or things together', zh: '组，群体' },
  country: { phonetic: '/ˈkʌntri/', en: 'a nation or territory', zh: '国家' },
  problem: { phonetic: '/ˈprɒbləm/', en: 'a matter difficult to deal with', zh: '问题' },
  fact: { phonetic: '/fækt/', en: 'a true piece of information', zh: '事实' },
  right: { phonetic: '/raɪt/', en: 'a moral or legal entitlement; correct', zh: '权利，正确，正确的' },
  study: { phonetic: '/ˈstʌdi/', en: 'the act of learning', zh: '学习，研究' },
  book: { phonetic: '/bʊk/', en: 'a written or printed work', zh: '书' },
  word: { phonetic: '/wɜːd/', en: 'a unit of language', zh: '单词' },
  business: { phonetic: '/ˈbɪznəs/', en: 'commercial activity', zh: '商业' },
  power: { phonetic: '/ˈpaʊər/', en: 'the ability to do something', zh: '力量，权力' },
  city: { phonetic: '/ˈsɪti/', en: 'a large town', zh: '城市' },
  market: { phonetic: '/ˈmɑːkɪt/', en: 'a regular gathering for trade', zh: '市场' },
  community: { phonetic: '/kəˈmjuːnəti/', en: 'a group of people living together', zh: '社区' },
  information: { phonetic: '/ˌɪnfəˈmeɪʃn/', en: 'data or knowledge', zh: '信息' },
  children: { phonetic: '/ˈtʃɪldrən/', en: 'young human beings', zh: '孩子们' },
  development: { phonetic: '/dɪˈveləpmənt/', en: 'the process of growing or improving', zh: '发展' },
  education: { phonetic: '/ˌedʒuˈkeɪʃn/', en: 'the process of teaching and learning', zh: '教育' },
  support: { phonetic: '/səˈpɔːt/', en: 'to bear all or part of the weight', zh: '支持' },
  research: { phonetic: '/rɪˈsɜːtʃ/', en: 'systematic investigation', zh: '研究' },
  difference: { phonetic: '/ˈdɪfrəns/', en: 'a way in which things are distinct', zh: '差异' },
  experience: { phonetic: '/ɪkˈspɪriəns/', en: 'practical contact with events', zh: '经验，体验' },
  result: { phonetic: '/rɪˈzʌlt/', en: 'a consequence or outcome', zh: '结果' },
  society: { phonetic: '/səˈsaɪəti/', en: 'the community of people', zh: '社会' },
  example: { phonetic: '/ɪɡˈzɑːmpl/', en: 'a thing characteristic of its kind', zh: '例子' },
  morning: { phonetic: '/ˈmɔːrnɪŋ/', en: 'the early part of the day', zh: '早晨' },
  moment: { phonetic: '/ˈməʊmənt/', en: 'a very brief period of time', zh: '瞬间' },
  story: { phonetic: '/ˈstɔːri/', en: 'a narrative of events', zh: '故事' },
  idea: { phonetic: '/aɪˈdɪə/', en: 'a thought or suggestion', zh: '想法' },
  data: { phonetic: '/ˈdeɪtə/', en: 'facts and statistics collected together', zh: '数据' },

  // Adjectives
  good: { phonetic: '/ɡʊd/', en: 'to be desired or approved of', zh: '好的' },
  new: { phonetic: '/njuː/', en: 'not existing before', zh: '新的' },
  first: { phonetic: '/fɜːst/', en: 'coming before all others in time', zh: '第一的' },
  last: { phonetic: '/lɑːst/', en: 'coming after all others', zh: '最后的' },
  long: { phonetic: '/lɒŋ/', en: 'measuring a great distance or duration', zh: '长的' },
  great: { phonetic: '/ɡreɪt/', en: 'of an extent considerably above average', zh: '伟大的，很' },
  little: { phonetic: '/ˈlɪtl/', en: 'small in size or amount', zh: '小的' },
  own: { phonetic: '/əʊn/', en: 'belonging to oneself', zh: '自己的' },
  other: { phonetic: '/ˈʌðər/', en: 'different; not the same', zh: '其他的' },
  old: { phonetic: '/əʊld/', en: 'having lived for a long time', zh: '老的，旧的' },
  big: { phonetic: '/bɪɡ/', en: 'of considerable size', zh: '大的' },
  high: { phonetic: '/haɪ/', en: 'of great vertical extent', zh: '高的' },
  different: { phonetic: '/ˈdɪfrənt/', en: 'not the same as another', zh: '不同的' },
  small: { phonetic: '/smɔːl/', en: 'little in size or degree', zh: '小的' },
  large: { phonetic: '/lɑːdʒ/', en: 'of considerable size', zh: '大的' },
  next: { phonetic: '/nekst/', en: 'coming immediately after', zh: '下一个的' },
  early: { phonetic: '/ˈɜːli/', en: 'before the expected time', zh: '早的' },
  young: { phonetic: '/jʌŋ/', en: 'having lived for a short time', zh: '年轻的' },
  important: { phonetic: '/ɪmˈpɔːtnt/', en: 'of great significance', zh: '重要的' },
  public: { phonetic: '/ˈpʌblɪk/', en: 'concerning the people as a whole', zh: '公共的' },
  bad: { phonetic: '/bæd/', en: 'of poor quality or low standard', zh: '坏的' },
  same: { phonetic: '/seɪm/', en: 'identical; not different', zh: '相同的' },
  able: { phonetic: '/ˈeɪbl/', en: 'having the power to do something', zh: '能够的' },
  possible: { phonetic: '/ˈpɒsəbl/', en: 'able to exist or happen', zh: '可能的' },
  true: { phonetic: '/truː/', en: 'in accordance with fact', zh: '真实的' },
  free: { phonetic: '/friː/', en: 'not confined or imprisoned', zh: '自由的' },
  full: { phonetic: '/fʊl/', en: 'containing as much as possible', zh: '满的' },
  sure: { phonetic: '/ʃʊər/', en: 'confident in what one thinks', zh: '确定的' },
  strong: { phonetic: '/strɒŋ/', en: 'having great physical power', zh: '强壮的' },
  special: { phonetic: '/ˈspeʃl/', en: 'better or greater than usual', zh: '特别的' },
  clear: { phonetic: '/klɪər/', en: 'easy to perceive or understand', zh: '清楚的' },
  hard: { phonetic: '/hɑːd/', en: 'solid and firm; difficult', zh: '困难的，硬的' },
  ready: { phonetic: '/ˈredi/', en: 'fully prepared for something', zh: '准备好的' },
  whole: { phonetic: '/həʊl/', en: 'complete; entire', zh: '全部的' },
  recent: { phonetic: '/ˈriːsnt/', en: 'having happened not long ago', zh: '最近的' },
  common: { phonetic: '/ˈkɒmən/', en: 'occurring or appearing frequently', zh: '常见的' },
  human: { phonetic: '/ˈhjuːmən/', en: 'relating to people', zh: '人类的' },
  natural: { phonetic: '/ˈnætʃrəl/', en: 'existing in nature', zh: '自然的' },
  certain: { phonetic: '/ˈsɜːtn/', en: 'known for sure; specific', zh: '确定的，某些' },
  available: { phonetic: '/əˈveɪləbl/', en: 'able to be used or obtained', zh: '可用的' },
  likely: { phonetic: '/ˈlaɪkli/', en: 'probable; expected', zh: '可能的' },
  simple: { phonetic: '/ˈsɪmpl/', en: 'easily understood; not complex', zh: '简单的' },

  // Adverbs
  not: { phonetic: '/nɒt/', en: 'used to form the negative', zh: '不' },
  so: { phonetic: '/səʊ/', en: 'to such a great extent', zh: '所以，如此' },
  very: { phonetic: '/ˈveri/', en: 'in a high degree', zh: '非常' },
  just: { phonetic: '/dʒʌst/', en: 'exactly; only', zh: '正好，仅仅' },
  also: { phonetic: '/ˈɔːlsəʊ/', en: 'in addition; too', zh: '也' },
  only: { phonetic: '/ˈəʊnli/', en: 'solely; exclusively', zh: '仅仅' },
  now: { phonetic: '/naʊ/', en: 'at the present moment', zh: '现在' },
  then: { phonetic: '/ðen/', en: 'at that time', zh: '那时' },
  here: { phonetic: '/hɪər/', en: 'in this place', zh: '这里' },
  there: { phonetic: '/ðeər/', en: 'in that place', zh: '那里' },
  well: { phonetic: '/wel/', en: 'in a good or satisfactory way', zh: '好地' },
  even: { phonetic: '/ˈiːvn/', en: 'used for emphasis', zh: '甚至' },
  still: { phonetic: '/stɪl/', en: 'up to this time', zh: '仍然' },
  always: { phonetic: '/ˈɔːlweɪz/', en: 'at all times', zh: '总是' },
  never: { phonetic: '/ˈnevər/', en: 'at no time', zh: '从不' },
  often: { phonetic: '/ˈɒfn/', en: 'frequently', zh: '经常' },
  sometimes: { phonetic: '/ˈsʌmtaɪmz/', en: 'occasionally', zh: '有时' },
  again: { phonetic: '/əˈɡen/', en: 'once more', zh: '再次' },
  too: { phonetic: '/tuː/', en: 'to a higher degree; also', zh: '太，也' },
  much: { phonetic: '/mʌtʃ/', en: 'to a great extent', zh: '很多' },
  really: { phonetic: '/ˈriːəli/', en: 'in actual fact; truly', zh: '真正地' },
  already: { phonetic: '/ɔːlˈredi/', en: 'before a specified time', zh: '已经' },
  quite: { phonetic: '/kwaɪt/', en: 'to a certain degree', zh: '相当' },
  however: { phonetic: '/haʊˈevər/', en: 'used to introduce a contrasting point', zh: '然而' },
  maybe: { phonetic: '/ˈmeɪbi/', en: 'perhaps; possibly', zh: '也许' },

  // Question words
  what: { phonetic: '/wɒt/', en: 'asking for information', zh: '什么' },
  when: { phonetic: '/wen/', en: 'at what time', zh: '什么时候' },
  where: { phonetic: '/weər/', en: 'in what place', zh: '哪里' },
  why: { phonetic: '/waɪ/', en: 'for what reason', zh: '为什么' },
  how: { phonetic: '/haʊ/', en: 'in what manner', zh: '如何' },
  who: { phonetic: '/huː/', en: 'what person', zh: '谁' },
  which: { phonetic: '/wɪtʃ/', en: 'asking for choice', zh: '哪一个' },

  // Pronouns
  i: { phonetic: '/aɪ/', en: 'oneself as a person', zh: '我' },
  you: { phonetic: '/juː/', en: 'the person being addressed', zh: '你' },
  he: { phonetic: '/hiː/', en: 'male person previously mentioned', zh: '他' },
  she: { phonetic: '/ʃiː/', en: 'female person previously mentioned', zh: '她' },
  it: { phonetic: '/ɪt/', en: 'a thing previously mentioned', zh: '它' },
  we: { phonetic: '/wiː/', en: 'oneself and others', zh: '我们' },
  they: { phonetic: '/ðeɪ/', en: 'people or things previously mentioned', zh: '他们' },
  me: { phonetic: '/miː/', en: 'objective case of "I"', zh: '我（宾格）' },
  him: { phonetic: '/hɪm/', en: 'objective case of "he"', zh: '他（宾格）' },
  her: { phonetic: '/hɜːr/', en: 'objective case of "she"', zh: '她（宾格）' },
  them: { phonetic: '/ðem/', en: 'objective case of "they"', zh: '他们（宾格）' },
  my: { phonetic: '/maɪ/', en: 'belonging to me', zh: '我的' },
  your: { phonetic: '/jɔːr/', en: 'belonging to you', zh: '你的' },
  his: { phonetic: '/hɪz/', en: 'belonging to him', zh: '他的' },
  its: { phonetic: '/ɪts/', en: 'belonging to it', zh: '它的' },
  our: { phonetic: '/ˈaʊər/', en: 'belonging to us', zh: '我们的' },
  their: { phonetic: '/ðeər/', en: 'belonging to them', zh: '他们的' },
  this: { phonetic: '/ðɪs/', en: 'referring to a specific thing here', zh: '这个' },
  that: { phonetic: '/ðæt/', en: 'referring to a specific thing there', zh: '那个' },
  these: { phonetic: '/ðiːz/', en: 'plural of "this"', zh: '这些' },
  those: { phonetic: '/ðəʊz/', en: 'plural of "that"', zh: '那些' },
  some: { phonetic: '/sʌm/', en: 'an unspecified number or amount', zh: '一些' },
  any: { phonetic: '/ˈeni/', en: 'one or some of a thing', zh: '任何' },
  all: { phonetic: '/ɔːl/', en: 'the whole quantity of', zh: '所有' },
  each: { phonetic: '/iːtʃ/', en: 'every one of two or more', zh: '每个' },
  every: { phonetic: '/ˈevri/', en: 'all of a group', zh: '每一个' },
  both: { phonetic: '/bəʊθ/', en: 'the two; the one as well as the other', zh: '两者' },
  no: { phonetic: '/nəʊ/', en: 'not any; not one', zh: '没有，不' },

  // Conjunctions
  and: { phonetic: '/ænd/', en: 'in addition; plus', zh: '和，而且' },
  but: { phonetic: '/bʌt/', en: 'used to introduce contrast', zh: '但是' },
  or: { phonetic: '/ɔːr/', en: 'used to link alternatives', zh: '或者' },
  because: { phonetic: '/bɪˈkɒz/', en: 'for the reason that', zh: '因为' },
  if: { phonetic: '/ɪf/', en: 'on the condition that', zh: '如果' },
  than: { phonetic: '/ðæn/', en: 'introducing a comparison', zh: '比' },
  while: { phonetic: '/waɪl/', en: 'during the time that', zh: '当…时' },
  although: { phonetic: '/ɔːlˈðəʊ/', en: 'in spite of the fact that', zh: '虽然' },
  since: { phonetic: '/sɪns/', en: 'from a past time until now', zh: '自从，因为' },
  unless: { phonetic: '/ʌnˈles/', en: 'except when', zh: '除非' },

  // Misc
  like: { phonetic: '/laɪk/', en: 'having similar qualities', zh: '像，喜欢' },
  about: { phonetic: '/əˈbaʊt/', en: 'on the subject of', zh: '关于' },
  into: { phonetic: '/ˈɪntə/', en: 'expressing movement to inside', zh: '进入' },
  over: { phonetic: '/ˈəʊvər/', en: 'extending directly above', zh: '在…上方' },
  after: { phonetic: '/ˈɑːftər/', en: 'in the time following', zh: '在…之后' },
  before: { phonetic: '/bɪˈfɔːr/', en: 'during the period preceding', zh: '在…之前' },
  between: { phonetic: '/bɪˈtwiːn/', en: 'in the space separating', zh: '在…之间' },
  under: { phonetic: '/ˈʌndər/', en: 'directly below', zh: '在…下面' },
  without: { phonetic: '/wɪðˈaʊt/', en: 'in the absence of', zh: '没有' },
  through: { phonetic: '/θruː/', en: 'moving in one side and out another', zh: '通过' },
  during: { phonetic: '/ˈdjʊərɪŋ/', en: 'throughout the course of', zh: '在…期间' },
  around: { phonetic: '/əˈraʊnd/', en: 'on every side of', zh: '在…周围' },
  against: { phonetic: '/əˈɡenst/', en: 'in opposition to', zh: '反对，靠着' },
  among: { phonetic: '/əˈmʌŋ/', en: 'in the midst of', zh: '在…之中' },
  across: { phonetic: '/əˈkrɒs/', en: 'from one side to the other', zh: '穿过' },
  behind: { phonetic: '/bɪˈhaɪnd/', en: 'at the back of', zh: '在…后面' },
  above: { phonetic: '/əˈbʌv/', en: 'in a higher position than', zh: '在…上面' },
  along: { phonetic: '/əˈlɒŋ/', en: 'moving on a surface or line', zh: '沿着' },
  within: { phonetic: '/wɪˈðɪn/', en: 'inside; not beyond', zh: '在…之内' },
  beyond: { phonetic: '/bɪˈjɒnd/', en: 'at the far side of', zh: '超出' },
  toward: { phonetic: '/təˈwɔːd/', en: 'in the direction of', zh: '朝向' },

  // More academic / interesting words
  significant: { phonetic: '/sɪɡˈnɪfɪkənt/', en: 'sufficiently great or important', zh: '显著的，重要的' },
  opportunity: { phonetic: '/ˌɒpəˈtjuːnəti/', en: 'a set of circumstances making something possible', zh: '机会' },
  environment: { phonetic: '/ɪnˈvaɪrənmənt/', en: 'the surroundings or conditions', zh: '环境' },
  individual: { phonetic: '/ˌɪndɪˈvɪdʒuəl/', en: 'a single person or thing', zh: '个人，个体' },
  particular: { phonetic: '/pəˈtɪkjʊlər/', en: 'used to single out an individual', zh: '特定的，特别的' },
  economic: { phonetic: '/ˌiːkəˈnɒmɪk/', en: 'relating to the economy', zh: '经济的' },
  political: { phonetic: '/pəˈlɪtɪkl/', en: 'relating to government or public affairs', zh: '政治的' },
  population: { phonetic: '/ˌpɒpjuˈleɪʃn/', en: 'all the inhabitants of a place', zh: '人口' },
  knowledge: { phonetic: '/ˈnɒlɪdʒ/', en: 'facts or information acquired through experience', zh: '知识' },
  technology: { phonetic: '/tekˈnɒlədʒi/', en: 'the application of scientific knowledge', zh: '技术' },
  strategy: { phonetic: '/ˈstrætədʒi/', en: 'a plan to achieve a long-term goal', zh: '策略' },
  analysis: { phonetic: '/əˈnæləsɪs/', en: 'detailed examination of something', zh: '分析' },
  identify: { phonetic: '/aɪˈdentɪfaɪ/', en: 'to establish who or what something is', zh: '识别，确认' },
  establish: { phonetic: '/ɪˈstæblɪʃ/', en: 'to set up on a firm basis', zh: '建立' },
  maintain: { phonetic: '/meɪnˈteɪn/', en: 'to keep in an existing state', zh: '维持' },
  increase: { phonetic: '/ɪnˈkriːs/', en: 'to become greater in size or amount', zh: '增加' },
  reduce: { phonetic: '/rɪˈdjuːs/', en: 'to make smaller or less', zh: '减少' },
  improve: { phonetic: '/ɪmˈpruːv/', en: 'to make or become better', zh: '改善' },
  concern: { phonetic: '/kənˈsɜːn/', en: 'a matter of interest or worry', zh: '关心，担忧' },
  affect: { phonetic: '/əˈfekt/', en: 'to influence or have an effect on', zh: '影响' },
  benefit: { phonetic: '/ˈbenɪfɪt/', en: 'an advantage or profit', zh: '好处' },
  challenge: { phonetic: '/ˈtʃælɪndʒ/', en: 'a task or situation that tests abilities', zh: '挑战' },
  current: { phonetic: '/ˈkʌrənt/', en: 'belonging to the present time', zh: '当前的' },
  effective: { phonetic: '/ɪˈfektɪv/', en: 'successful in producing a desired result', zh: '有效的' },
  emerge: { phonetic: '/ɪˈmɜːdʒ/', en: 'to come out into view', zh: '出现' },
  focus: { phonetic: '/ˈfəʊkəs/', en: 'the center of interest or activity', zh: '焦点，集中' },
  global: { phonetic: '/ˈɡləʊbl/', en: 'relating to the whole world', zh: '全球的' },
  impact: { phonetic: '/ˈɪmpækt/', en: 'a marked effect or influence', zh: '影响' },
  issue: { phonetic: '/ˈɪʃuː/', en: 'an important topic or problem', zh: '问题，议题' },
  major: { phonetic: '/ˈmeɪdʒər/', en: 'important; serious; significant', zh: '主要的' },
  process: { phonetic: '/ˈprəʊses/', en: 'a series of actions or steps', zh: '过程' },
  require: { phonetic: '/rɪˈkwaɪər/', en: 'to need for a particular purpose', zh: '需要' },
  resource: { phonetic: '/rɪˈsɔːs/', en: 'a supply that can be drawn upon', zh: '资源' },
  response: { phonetic: '/rɪˈspɒns/', en: 'an answer or reaction', zh: '回应' },
  source: { phonetic: '/sɔːs/', en: 'origin or starting point', zh: '来源' },
  structure: { phonetic: '/ˈstrʌktʃər/', en: 'the arrangement of parts', zh: '结构' },
  traditional: { phonetic: '/trəˈdɪʃənl/', en: 'existing in or as part of a tradition', zh: '传统的' },
  variety: { phonetic: '/vəˈraɪəti/', en: 'diversity; a number of different things', zh: '多样性' },
  achieve: { phonetic: '/əˈtʃiːv/', en: 'to reach or attain by effort', zh: '达到，实现' },
  recognize: { phonetic: '/ˈrekəɡnaɪz/', en: 'to identify from previous encounters', zh: '认出，承认' },
  indicate: { phonetic: '/ˈɪndɪkeɪt/', en: 'to point out or show', zh: '表明' },
  generate: { phonetic: '/ˈdʒenəreɪt/', en: 'to produce or create', zh: '产生' },
  define: { phonetic: '/dɪˈfaɪn/', en: 'to state the meaning of', zh: '定义' },
  function: { phonetic: '/ˈfʌŋkʃn/', en: 'an activity or purpose natural to something', zh: '功能，函数' },
  theory: { phonetic: '/ˈθɪəri/', en: 'a system of ideas explaining something', zh: '理论' },
  concept: { phonetic: '/ˈkɒnsept/', en: 'an abstract idea or notion', zh: '概念' },
  context: { phonetic: '/ˈkɒntekst/', en: 'the circumstances surrounding an event', zh: '上下文，背景' },
};

// ─── Dictionary API placeholder ─────────────────────
export async function fetchDictionaryAPI(word: string): Promise<DictEntry | null> {
  // TODO: Replace with real API call (e.g. Merriam-Webster, Oxford, Youdao)
  // const res = await fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${word}`);
  // const data = await res.json();
  // return { phonetic: data[0]?.phonetic, en: data[0]?.meanings[0]?.definitions[0]?.definition, zh: '...' };

  // Fallback to local mock
  const MOCK_LOOKUP: Record<string, DictEntry> = MOCK_DICT;

  // Try exact match first
  let entry = MOCK_LOOKUP[word.toLowerCase()];
  if (entry) return entry;

  // Try stripping common suffixes
  const suffixes = ['s', 'es', 'ies', 'ed', 'd', 'ied', 'ing', 'ly', 'er', 'est', 'tion', 'ment'];
  for (const suffix of suffixes) {
    if (word.toLowerCase().endsWith(suffix)) {
      const root = word.toLowerCase().slice(0, -suffix.length);
      entry = MOCK_LOOKUP[root];
      if (entry) return entry;
      // try adding 'e' back (e.g. "baking" → "bake" → "bak" + "e")
      if (MOCK_LOOKUP[root + 'e']) return MOCK_LOOKUP[root + 'e'];
    }
  }

  return null;
}

// ─── Speech ──────────────────────────────────────────
function speakWord(word: string) {
  if (typeof window === 'undefined') return;
  window.speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(word);
  u.rate = 0.85;
  u.lang = 'en-US';
  const voices = window.speechSynthesis.getVoices();
  const voice = voices.find((v) => v.lang.startsWith('en'));
  if (voice) u.voice = voice;
  window.speechSynthesis.speak(u);
}

// ─── Sentence splitting ─────────────────────────────
function splitSentences(text: string): string[] {
  const t = text.replace(/\s+/g, ' ').trim();
  if (!t) return [];

  // Protect abbreviations
  const abbrMarkers: { marker: string; original: string }[] = [];
  let processed = t;

  const patterns: [RegExp, string, number][] = [
    [/\b(Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|Sgt|Capt|Lt|Gen|Col|Sra|Srta)\./gi, '$1\x01', 0],
    [/\b(etc|vs|dept|est|approx|govt|Co|Inc|Ltd|Corp|Plc|LLC)\./gi, '$1\x02', 1],
    [/\b([A-Z])\.([A-Z])\./g, '$1\x03A', 2],
    [/\b(a\.m|p\.m)\./gi, '$1\x04', 3],
    [/\b(Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\./gi, '$1\x05', 4],
  ];

  patterns.forEach(([regex, replacement, idx]) => {
    const marker = String.fromCharCode(0x01 + idx);
    processed = processed.replace(regex, (...args) => {
      const match = args[0] as string;
      abbrMarkers.push({ marker, original: match });
      return match.replace(/\./g, marker);
    });
  });

  // Split at sentence boundaries
  const raw = processed.split(/(?<=[.!?])\s+(?=["'‘’“"(（\d]*[A-Z0-9])/);

  return raw
    .map((s: string) => {
      let restored = s;
      for (const { marker, original } of abbrMarkers) {
        restored = restored.replace(new RegExp(marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '.');
      }
      restored = restored.replace(/[\x01-\x05]/g, '.');
      return restored.trim();
    })
    .filter(Boolean);
}

// ─── Word tokenization ───────────────────────────────
function tokenizeWords(sentence: string): WordToken[] {
  const parts = sentence.match(/\S+/g) || [];
  let wordIdx = 0;
  const sentId = Math.random().toString(36).slice(2, 8);

  return parts.map((token) => {
    const id = `${sentId}_${wordIdx++}`;
    const clean = token.replace(/^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$/g, '');
    return { display: token, clean, id };
  });
}

function parseArticle(text: string): SentenceData[] {
  const sentences = splitSentences(text);
  return sentences.map((s) => ({
    words: tokenizeWords(s),
    id: Math.random().toString(36).slice(2, 10),
  }));
}

// ─── Component ───────────────────────────────────────
interface ArticleParserProps {
  onWordClick?: (word: string) => void;
  placeholder?: string;
  vocab: Set<string>;
  onToggleVocab: (word: string) => void;
}

export default function ArticleParser({ onWordClick, placeholder, vocab, onToggleVocab }: ArticleParserProps) {
  const [text, setText] = useState('');
  const [parsed, setParsed] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [activeWord, setActiveWord] = useState<WordToken | null>(null);
  const [dictData, setDictData] = useState<DictApiResponse | null>(null);
  const [isDictLoading, setIsDictLoading] = useState(false);
  const [playingSentenceId, setPlayingSentenceId] = useState<string | null>(null);
  const [testMode, setTestMode] = useState(false);
  const [correctTestIds, setCorrectTestIds] = useState<Set<string>>(new Set());
  const [flashIds, setFlashIds] = useState<Set<string>>(new Set());
  const tooltipRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const ttsAbortRef = useRef(false);
  const testInputRefs = useRef<Map<string, HTMLInputElement>>(new Map());

  const sentences = useMemo(() => {
    if (!parsed || !text.trim()) return [];
    return parseArticle(text);
  }, [text, parsed]);

  // ── Word click handler ──
  const handleWordClick = useCallback(
    async (e: React.MouseEvent, word: WordToken) => {
      e.stopPropagation();
      if (!word.clean) return;

      onWordClick?.(word.clean);

      // Stop sentence playback when user clicks a word
      if (playingSentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
      }

      // Show tooltip immediately with loading state
      setActiveWord(word);
      setDictData(null);
      setIsDictLoading(true);

      // Fetch dictionary definition
      try {
        const res = await fetch(`/api/dict?word=${encodeURIComponent(word.clean)}`);
        if (res.ok) {
          const data = await res.json();
          if (data.error) {
            setDictData(null);
          } else {
            setDictData(data as DictApiResponse);
          }
        } else {
          setDictData(null);
        }
      } catch {
        setDictData(null);
      } finally {
        setIsDictLoading(false);
      }
    },
    [onWordClick, playingSentenceId],
  );

  // ── Close tooltip on outside click ──
  useEffect(() => {
    if (!activeWord) return;

    const handleClickOutside = (e: MouseEvent) => {
      if (tooltipRef.current && !tooltipRef.current.contains(e.target as Node)) {
        setActiveWord(null);
        setDictData(null);
    
      }
    };

    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setActiveWord(null);
        setDictData(null);
    
      }
    };

    const handleScroll = () => {
      setActiveWord(null);
      setDictData(null);
  
    };

    const id = setTimeout(() => {
      document.addEventListener('click', handleClickOutside);
      document.addEventListener('keydown', handleEscape);
      window.addEventListener('scroll', handleScroll, true);
    }, 0);

    return () => {
      clearTimeout(id);
      document.removeEventListener('click', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', handleScroll, true);
    };
  }, [activeWord]);

  // ── Parse handler ──
  const handleParse = useCallback(() => {
    // Read directly from DOM to avoid stale React state during hydration
    const currentText = textareaRef.current?.value ?? text;
    if (!currentText.trim()) return;
    if (!parsed || currentText !== text) {
      setText(currentText);
    }
    ttsAbortRef.current = true;
    stopPlayback();
    setParsed(true);
    setActiveWord(null);
    setDictData(null);

    setPlayingSentenceId(null);
    if (testMode) {
      setTestMode(false);
      setCorrectTestIds(new Set());
      setFlashIds(new Set());
    }
  }, [text, testMode, parsed]);

  // ── Text change handler ──
  const handleTextChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      setText(e.target.value);
      if (parsed) setParsed(false);
      setActiveWord(null);
      setDictData(null);
      if (playingSentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
      }
    },
    [parsed, playingSentenceId],
  );

  // ── Keyboard shortcut ──
  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        handleParse();
      }
    },
    [handleParse],
  );

  // ── Close dictionary ──
  const closeDict = useCallback(() => {
    setActiveWord(null);
    setDictData(null);

  }, []);

  // ── Sentence TTS playback ──
  const handleSentencePlay = useCallback(
    async (sentenceId: string, text: string) => {
      // If already playing this sentence, stop it
      if (playingSentenceId === sentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
        return;
      }

      // Stop any current playback and start new one
      ttsAbortRef.current = false;
      stopPlayback();
      setPlayingSentenceId(sentenceId);
      setActiveWord(null);
      setDictData(null);
  

      try {
        await speakSentence(text);
        if (!ttsAbortRef.current) {
          setPlayingSentenceId(null);
        }
      } catch (err) {
        if (!ttsAbortRef.current) {
          console.error('TTS playback error:', err);
          setPlayingSentenceId(null);
        }
      }
    },
    [playingSentenceId],
  );

  // Clean up TTS on unmount
  useEffect(() => {
    return () => {
      ttsAbortRef.current = true;
      stopPlayback();
    };
  }, []);

  // ── Test mode ──
  const toggleTestMode = useCallback(() => {
    setTestMode((prev) => {
      const next = !prev;
      if (prev) {
        // Exiting test mode — reset answers
        setCorrectTestIds(new Set());
        setFlashIds(new Set());
      } else {
        // Entering test mode — stop playback, close dictionary
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
        setActiveWord(null);
        setDictData(null);
    
      }
      return next;
    });
  }, []);

  const handleTestKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>, wordId: string, correctAnswer: string) => {
      if (e.key !== 'Enter') return;
      const input = e.currentTarget;
      const typed = input.value.trim();

      if (typed.toLowerCase() === correctAnswer.toLowerCase()) {
        setCorrectTestIds((prev) => {
          const next = new Set(prev);
          next.add(wordId);
          return next;
        });
        // Auto-advance to the next unfilled input by DOM order
        requestAnimationFrame(() => {
          const inputs = document.querySelectorAll<HTMLInputElement>('input[data-word-id]');
          const idx = Array.from(inputs).findIndex((el) => el.dataset.wordId === wordId);
          const next = Array.from(inputs).slice(idx + 1).find((el) => !el.value);
          if (next) next.focus();
        });
      } else {
        // Flash red
        setFlashIds((prev) => new Set(prev).add(wordId));
        setTimeout(() => {
          setFlashIds((prev) => {
            const next = new Set(prev);
            next.delete(wordId);
            return next;
          });
          // Re-focus
          const inp = testInputRefs.current.get(wordId);
          if (inp) { inp.value = ''; inp.focus(); }
        }, 400);
      }
    },
    [],
  );

  // ── Export to Word ──
  const [exporting, setExporting] = useState(false);

  const exportToWord = useCallback(async () => {
    if (vocab.size === 0 || exporting) return;
    setExporting(true);
    try {
      const vocabArray = Array.from(vocab);

      // One paragraph per word
      const wordParagraphs = vocabArray.map(
        (word) =>
          new Paragraph({
            spacing: { before: 120, after: 120 },
            children: [new TextRun({ text: word, size: 24, bold: true })],
          }),
      );

      // Build the document
      const doc = new Document({
        title: '雅思外刊阅读精读生词本',
        creator: 'English Learning App',
        styles: {
          default: {
            document: {
              run: { font: 'Arial', size: 22 },
              paragraph: { spacing: { after: 100 } },
            },
          },
        },
        sections: [
          {
            children: [
              new Paragraph({
                text: '雅思外刊阅读精读生词本',
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: { after: 600, before: 200 },
              }),
              ...wordParagraphs,
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'IELTS_My_Vocabulary.docx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting Word document:', err);
    } finally {
      setExporting(false);
    }
  }, [vocab, exporting]);

  // ── Derived state ──
  const hasText = text.trim().length > 0;
  const parsedCount = sentences.reduce((sum, s) => sum + s.words.length, 0);

  // ── Render ──
  return (
    <div ref={containerRef} className="mx-auto max-w-3xl">
      {/* ── Input area ── */}
      <div className="mb-6 space-y-3">
        <textarea
          value={text}
          onChange={handleTextChange}
          onKeyDown={handleKeyDown}
          placeholder={placeholder ?? 'Paste an English article here...'}
          rows={6}
          className="w-full resize-y rounded-lg border border-gray-300 bg-white p-4 text-base leading-relaxed text-gray-900 shadow-sm transition-colors placeholder:text-gray-400 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500/20"
        />

        <div className="flex items-center justify-between">
          <span className="text-sm text-gray-500">
            {text.length > 0 && `${text.length} characters`}
          </span>

          <button
            onClick={handleParse}
            className="inline-flex items-center gap-2 rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500/40 touch-manipulation min-h-[44px]"
          >
            <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
            Parse
          </button>
        </div>

        {parsed && (
          <div className="flex items-center justify-between">
            <p className="text-xs text-gray-400">
              Ctrl / Cmd + Enter &middot; {parsedCount} words across {sentences.length} sentences
              {vocab.size > 0 && ` · ${vocab.size} saved`}
            </p>
            {vocab.size > 0 && (
              <div className="flex items-center gap-2">
                <button
                  onClick={exportToWord}
                  disabled={exporting}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm shadow-violet-300 transition-all hover:from-violet-700 hover:to-indigo-700 hover:shadow-md hover:shadow-violet-400 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  {exporting ? 'Generating...' : 'Export Word'}
                </button>
                <button
                  onClick={toggleTestMode}
                  className={`
                    inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all
                    ${
                      testMode
                        ? 'bg-indigo-100 text-indigo-700 ring-1 ring-indigo-300 hover:bg-indigo-200'
                        : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }
                  `}
                >
                  <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    {testMode
                      ? <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                      : <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    }
                  </svg>
                  {testMode ? 'Exit Test Mode' : 'Dictation Test'}
                </button>
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── Empty state ── */}
      {!parsed && (
        <div className="flex flex-col items-center justify-center rounded-xl border-2 border-dashed border-gray-200 py-16 text-center text-gray-400">
          <svg className="mb-3 h-10 w-10" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
          </svg>
          <p className="text-sm">Paste an article above and click Parse to get started</p>
        </div>
      )}

      {/* ── Parse failure ── */}
      {parsed && sentences.length === 0 && hasText && (
        <div className="rounded-xl border border-yellow-200 bg-yellow-50 px-4 py-6 text-center text-sm text-yellow-700">
          Could not split the text into sentences. Try pasting a different article.
        </div>
      )}

      {/* ── Test mode banner ── */}
      {testMode && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border border-indigo-200 bg-indigo-50 px-4 py-2.5 text-sm text-indigo-700">
          <svg className="h-4 w-4 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
          <span>
            <strong>Dictation Test Mode</strong> — Type the correct spelling of each word and press <kbd className="rounded bg-indigo-200 px-1 py-0.5 font-mono text-xs">Enter</kbd> to check.
          </span>
        </div>
      )}

      {/* ── Parse result ── */}
      {parsed && sentences.length > 0 && (
        <div className={`space-y-3 ${testMode ? 'rounded-xl border border-indigo-100 bg-white/80 p-4' : ''}`}>
          {sentences.map((sentence) => {
            const isPlaying = playingSentenceId === sentence.id;
            const sentenceText = sentence.words
              .map((w) => w.display)
              .join(' ')
              .trim();

            return (
              <div
                key={sentence.id}
                className={`
                  group relative flex items-start gap-2 rounded-xl px-3 py-2
                  transition-all duration-300
                  ${
                    isPlaying
                      ? 'bg-blue-50 shadow-sm ring-1 ring-blue-200'
                      : 'hover:bg-gray-50/60'
                  }
                `}
              >
                {/* ── Sentence play button ── */}
                {typeof window !== 'undefined' && window.speechSynthesis && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleSentencePlay(sentence.id, sentenceText);
                    }}
                    aria-label={isPlaying ? 'Stop' : 'Play sentence'}
                    title={isPlaying ? 'Stop' : 'Play sentence'}
                    className={`
                      mt-1 flex-shrink-0 rounded-lg p-1.5 transition-colors
                      focus:outline-none focus:ring-2 focus:ring-blue-300
                      ${
                        isPlaying
                          ? 'bg-blue-100 text-blue-600 hover:bg-blue-200'
                          : 'text-gray-400 opacity-0 group-hover:opacity-100 hover:text-blue-500 hover:opacity-100'
                      }
                    `}
                  >
                    {isPlaying ? (
                      /* Stop icon */
                      <svg className="h-4 w-4" viewBox="0 0 24 24" fill="currentColor">
                        <rect x="6" y="6" width="12" height="12" rx="1" />
                      </svg>
                    ) : (
                      /* Play icon */
                      <svg className="h-4 w-4" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M8 5.14v13.72a1 1 0 001.5.86l11-6.86a1 1 0 000-1.72l-11-6.86A1 1 0 008 5.14z" />
                      </svg>
                    )}
                  </button>
                )}

                {/* ── Words ── */}
                <p className="flex-1 leading-[2.2]">
                  {sentence.words.map((word) => {
                    const isSaved = vocab.has(word.clean);

                    // ── Test mode: render input for vocab words ──
                    if (testMode && isSaved) {
                      const isCorrect = correctTestIds.has(word.id);
                      const isFlashing = flashIds.has(word.id);

                      if (isCorrect) {
                        return (
                          <span
                            key={word.id}
                            className="inline-block rounded px-0.5 text-base font-medium leading-[2.2] text-green-600"
                          >
                            {word.clean}{' '}
                          </span>
                        );
                      }

                      return (
                        <input
                          key={word.id}
                          ref={(el) => {
                            if (el) testInputRefs.current.set(word.id, el);
                            else testInputRefs.current.delete(word.id);
                          }}
                          type="text"
                          defaultValue=""
                          spellCheck={false}
                          autoComplete="off"
                          data-word-id={word.id}
                          placeholder={' '.repeat(Math.max(word.clean.length - 1, 1))}
                          data-word={word.clean}
                          style={{
                            width: `calc(${word.clean.length}ch + 1.2rem)`,
                            minWidth: '4rem',
                          }}
                          onKeyDown={(e) => handleTestKeyDown(e, word.id, word.clean)}
                          className={`
                            mx-0.5 inline-block rounded border-b-2 bg-transparent
                            px-1.5 py-0 text-center text-base
                            outline-none transition-all duration-200
                            ${
                              isFlashing
                                ? 'border-red-400 bg-red-50'
                                : 'border-gray-300 focus:border-indigo-400'
                            }
                          `}
                        />
                      );
                    }

                    // ── Normal mode: clickable word ──
                    return (
                      <span
                        key={word.id}
                        onClick={(e) => handleWordClick(e, word)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault();
                            const clickEvent = { ...e, target: e.currentTarget } as unknown as React.MouseEvent;
                            handleWordClick(clickEvent, word);
                          }
                        }}
                        tabIndex={testMode ? -1 : 0}
                        role="button"
                        aria-label={`Word: ${word.clean}`}
                        className={`
                          inline-block cursor-pointer rounded px-0.5 text-base leading-[2.2] transition-colors
                          focus:outline-none focus:ring-2 focus:ring-yellow-300/60
                          ${isSaved ? 'text-red-600 hover:bg-red-100' : 'text-gray-800 hover:bg-yellow-200'}
                        `}
                      >
                        {word.display}{' '}
                      </span>
                    );
                  })}
                </p>
              </div>
            );
          })}
        </div>
      )}

      {/* ── Dictionary Popover ── */}
      {activeWord && (
        <div className="fixed top-[15%] left-1/2 -translate-x-1/2 z-[9999]">
          <Draggable nodeRef={tooltipRef} handle=".drag-handle">
            <div
              ref={tooltipRef}
              role="dialog"
              aria-label={`Dictionary: ${activeWord.clean}`}
              className="flex flex-col resize overflow-hidden min-w-[300px] min-h-[250px] max-w-[90vw] max-h-[90vh] pb-8 animate-in fade-in slide-in-from-top-2 rounded-xl border border-gray-200 bg-white shadow-xl"
            >
            {/* Drag handle header */}
            <div className="drag-handle cursor-move flex-shrink-0 flex items-center justify-between border-b border-gray-100 bg-gray-50 px-4 py-2.5 rounded-t-xl">
              <div className="flex items-center gap-2 min-w-0">
                <svg className="h-4 w-4 flex-shrink-0 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 8h16M4 16h16" />
                </svg>
                <span className="text-sm font-bold text-gray-800 truncate">{activeWord.clean}</span>
              </div>
              <div className="flex items-center gap-1">
                {/* Speak button */}
                <button
                  onClick={(e) => { e.stopPropagation(); speakWord(activeWord.clean); }}
                  aria-label="Pronounce"
                  className="rounded-lg p-1.5 text-gray-500 transition-colors hover:bg-gray-100 hover:text-gray-700"
                >
                  <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M15.536 8.464a5 5 0 010 7.072M17.95 6.05a8 8 0 010 11.9M6.5 8.8l4.3-4.3a1 1 0 011.7.7v13.6a1 1 0 01-1.7.7l-4.3-4.3H4a1 1 0 01-1-1v-4.4a1 1 0 011-1h2.5z" />
                  </svg>
                </button>
                {/* Close button */}
                <button
                  onClick={(e) => { e.stopPropagation(); closeDict(); }}
                  aria-label="Close"
                  className="rounded-lg p-1.5 text-gray-400 transition-colors hover:bg-gray-100 hover:text-gray-600"
                >
                  <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                  </svg>
                </button>
              </div>
            </div>

            {/* Body */}
            <div className="flex min-h-0 flex-1 flex-col px-4 py-3">
              {isDictLoading ? (
                <div className="flex flex-col items-center justify-center py-8 gap-2">
                  <svg className="h-5 w-5 animate-spin text-indigo-400" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                  </svg>
                  <span className="text-xs text-gray-400">Looking up &ldquo;{activeWord.clean}&rdquo;...</span>
                </div>
              ) : dictData && 'html' in dictData ? (
                <div
                  className="oxford-dict oxford-dict-content text-sm p-4 h-full w-full overflow-y-auto"
                  dangerouslySetInnerHTML={{ __html: dictData.html }}
                />
              ) : dictData ? (
                <div className="flex min-h-0 flex-1 flex-col overflow-y-auto space-y-3">
                  {/* Phonetic + Part of Speech */}
                  {(dictData.phonetic || dictData.partOfSpeech) && (
                    <div className="flex items-center gap-2 text-sm">
                      {dictData.phonetic && (
                        <span className="text-gray-500 font-mono">{dictData.phonetic}</span>
                      )}
                      {dictData.partOfSpeech && (
                        <span className="inline-block rounded bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-600 italic">
                          {dictData.partOfSpeech}
                        </span>
                      )}
                    </div>
                  )}

                  {/* English definitions */}
                  {dictData.enDefinitions.length > 0 && (
                    <div>
                      <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">English</span>
                      <ol className="mt-1 space-y-1.5 list-decimal list-inside">
                        {dictData.enDefinitions.map((def, i) => (
                          <li key={i} className="text-sm text-gray-700 leading-relaxed">{def}</li>
                        ))}
                      </ol>
                    </div>
                  )}

                  {/* Chinese definition */}
                  {dictData.zhDefinition && (
                    <div>
                      <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">中文</span>
                      <p className="mt-1 text-sm text-gray-700">{dictData.zhDefinition}</p>
                    </div>
                  )}

                  {/* Source */}
                  {dictData.source && (
                    <p className="text-[10px] text-gray-300 text-right mt-2">{dictData.source}</p>
                  )}
                </div>
              ) : (
                <p className="py-4 text-center text-sm text-gray-400">
                  No definition found for &ldquo;{activeWord.clean}&rdquo;
                </p>
              )}
            </div>

            {/* Footer */}
            {!isDictLoading && (
              <div className="flex-shrink-0 border-t border-gray-100 px-4 py-2.5">
                <button
                  onClick={(e) => { e.stopPropagation(); onToggleVocab(activeWord.clean); }}
                  className={`flex w-full items-center justify-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium transition-colors ${
                    vocab.has(activeWord.clean)
                      ? 'bg-red-50 text-red-600 hover:bg-red-100'
                      : 'bg-gray-50 text-gray-600 hover:bg-gray-100 hover:text-gray-800'
                  }`}
                >
                  <svg className="h-4 w-4" fill={vocab.has(activeWord.clean) ? 'currentColor' : 'none'} viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M5 5a2 2 0 012-2h10a2 2 0 012 2v16l-7-3.5L5 21V5z" />
                  </svg>
                  {vocab.has(activeWord.clean) ? 'Remove from vocabulary' : 'Add to vocabulary'}
                </button>
              </div>
            )}

            {/* Resize handle indicator */}
            <div className="pointer-events-none absolute bottom-0 right-0 flex items-end justify-end p-1 text-gray-300">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" className="drop-shadow-sm">
                <line x1="11" y1="16" x2="16" y2="11" />
                <line x1="6" y1="16" x2="16" y2="6" />
                <line x1="1" y1="16" x2="16" y2="1" />
              </svg>
            </div>
          </div>
          </Draggable>
        </div>
      )}
    </div>
  );
}
