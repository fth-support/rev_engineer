import { useState } from 'react'
import { Link } from 'react-router-dom'
import {
  ScanLine, Smartphone, Users, CreditCard, Receipt, Wifi, Globe, Boxes,
  ShieldCheck, ArrowRight, ExternalLink, Sparkles, Network, Database, Server,
  Loader, Tag, LayoutGrid, Plug, ListChecks, Wrench, CircleDot,
  SlidersHorizontal, Palette, Layers,
} from 'lucide-react'
import './pages.css'
import './sco.css'

// Served from knowledge-pool/public/sco-kiosk/ — base-relative for GitHub Pages.
const DEMO = `${import.meta.env.BASE_URL}sco-kiosk/index.html`
const CONFIG = `${import.meta.env.BASE_URL}sco-kiosk/config.html`

// The real state machine from the prototype: idle → member-first → scan → pay → processing → receipt.
const FLOW = [
  { n: '01', icon: Smartphone, title: 'Idle', th: 'แตะเพื่อเริ่ม', desc: 'Attract loop: promo carousel, OPEN 24H, language ไทย/EN/中文, call-staff lane LED.' },
  { n: '02', icon: Users, title: 'Member', th: 'สมาชิก', desc: 'Member-first: phone / app QR / co-brand card / coupon — or skip as guest (PDPA-masked).' },
  { n: '03', icon: ScanLine, title: 'Scan & Cart', th: 'สแกนสินค้า', desc: 'Scan, qty ±, delete, group repeats, promo & member pricing, running total.' },
  { n: '04', icon: CreditCard, title: 'Payment', th: 'ชำระเงิน', desc: 'Order summary + discounts; pick method (card / QR / Alipay / WeChat) via BBL EDC.' },
  { n: '05', icon: Loader, title: 'Processing', th: 'กำลังชำระ', desc: 'EDC-mirrored guidance; retry on failure with no money taken.' },
  { n: '06', icon: Receipt, title: 'Receipt', th: 'ใบเสร็จ', desc: 'None / print ABB / e-receipt (→ Foodland landing); auto-reset to idle.' },
]

const SCREENS = ['Idle / Attract', 'Member-first', 'Scan & Cart', 'Member & Coupon', 'Payment method', 'Processing', 'Receipt', 'Staff (PIN)', 'Out of service']
const OVERLAYS = ['Weighed produce', 'Price entry', 'Attendant / age', 'Unknown barcode', 'Call staff', 'Manual keypad', 'Inactivity timeout', 'Coupon code']

const CAPS = [
  { icon: Globe, tone: 'var(--accent)', title: 'Trilingual UI', text: 'ไทย / English / 中文 on every screen; voice guidance + more languages on the roadmap.' },
  { icon: Wifi, tone: 'var(--success)', title: 'Offline-first', text: 'Full checkout works offline; the sale queues locally and syncs when back online.' },
  { icon: CreditCard, tone: 'var(--accent-2)', title: 'EDC payments', text: 'Card, QR PromptPay / QRCS, Alipay, WeChat — the BBL EDC prints the slip.' },
  { icon: Receipt, tone: 'var(--accent)', title: 'Thai fiscal', text: 'Registered POS Tax ID, abbreviated tax invoice (ABB) and closing report.' },
  { icon: Boxes, tone: 'var(--success)', title: 'Edge cases', text: 'Weighed produce, age-restricted approval, price-required & unknown barcode.' },
  { icon: ShieldCheck, tone: 'var(--accent-2)', title: 'Staff mode', text: 'PIN-gated reprint ABB, close shift, close program, out-of-service.' },
]

// Promotion families (from the Fujitsu BBY master — to confirm final scope with Foodland).
const PROMO_SELF = ['Outright / time (P001)', 'Member price (P010)', 'Deal & free (D/F)', 'Auto trandis (P003)', 'SKU point (P012)', 'Lower price (P011)', 'Consignment (P002)']
const PROMO_OPS = ['Coupon redeem (P005)', 'Cooperate (P006/P007)', 'On-top (P009)', 'Transaction discount (P013)', 'Line markdown (P015)', 'Manual (M001–M006)', 'Installment (I001)']

// Fujitsu POS API touch-points referenced in the kickoff.
const APIS = [
  { id: '#0', label: 'Display authorize' },
  { id: '#1', label: 'Customer lookup' },
  { id: '#2', label: 'Item / promotion' },
  { id: '#7', label: 'Coupon validate' },
  { id: '#10', label: 'Send transaction' },
  { id: '#11', label: 'Print ABB' },
  { id: '#13', label: 'Logout' },
  { id: '#14', label: 'Closing report' },
]

const ARCH = [
  { icon: Database, label: 'Customer ERP' },
  { icon: Server, label: 'POS Server (HQ)' },
  { icon: Server, label: 'POS Store Server' },
  { icon: Network, label: 'POS API (Fujitsu)' },
  { icon: ScanLine, label: 'SCO lane' },
]

const MOCKED = ['No backend — catalog / members / prices / points hardcoded', 'No hardware — scanner, scale, EDC, printer, camera simulated', 'Payments are a timed simulation (force success/fail)', 'ABB / receipt is visual only', '"Reach mode" a11y toggle is a stub', 'Voice + JP/AR languages are spec, not built']
const NEXT = ['Decide platform — keep DC-HTML or port to React/TS (same state shape)', 'Extract data behind a service layer / API (no hardcoded catalog)', 'Define the SCO↔POS API contract (item, promo/quota, coupon, member, submit, ABB)', 'Hardware abstraction — scanner, scale, EDC, printer, camera', 'Finish or cut reach mode; build voice + remaining languages', 'Harden edge cases — partial pay, void/refund, network loss mid-payment, reprint, shift reconciliation']

function SCO() {
  const [loaded, setLoaded] = useState(false)
  const [cfgLoaded, setCfgLoaded] = useState(false)

  return (
    <div className="sco">
      {/* hero */}
      <section className="sco-hero glass">
        <div className="sco-hero__main">
          <span className="sco-badge"><Sparkles size={14} /> PROJECT SCO · WHITE-LABEL</span>
          <h1>One engine, any brand</h1>
          <p>
            A <strong>modular, white-label</strong> self-checkout platform. One standard engine that any
            retailer can <strong>recolor, rebrand and trim</strong> — integrated through a single
            <strong> POS API</strong>, so the lane is just one more POS terminal that pulls SKU, promotions
            &amp; loyalty and writes sales back. The live demo is configured for <strong>Foodland</strong>,
            the reference tenant.
          </p>
          <div className="sco-hero__cta">
            <a className="btn primary" href={DEMO} target="_blank" rel="noreferrer">
              <ScanLine size={16} /> Open kiosk demo <ExternalLink size={14} />
            </a>
            <Link to="/flows" className="btn ghost">See the POS foundation <ArrowRight size={15} /></Link>
          </div>
        </div>
        <div className="sco-hero__chips">
          <span className="chip">1080×1920 kiosk</span>
          <span className="chip">ไทย / EN / 中文</span>
          <span className="chip">Offline-first</span>
          <span className="chip">EDC · PromptPay · Alipay</span>
        </div>
      </section>

      {/* live demo */}
      <section className="sco-demo glass">
        <div className="sco-section-head">
          <h2><ScanLine size={20} color="var(--accent)" /> Interactive prototype</h2>
          <a className="sco-open" href={DEMO} target="_blank" rel="noreferrer">Open fullscreen <ExternalLink size={14} /></a>
        </div>
        <p className="sco-demo__hint">Tap through the real shopper loop, or use the left control panel to jump to any state / overlay.</p>
        <div className="sco-frame">
          {!loaded && <div className="sco-frame__loading">Loading kiosk…</div>}
          <iframe
            title="Foodland SCO kiosk demo"
            src={DEMO}
            loading="lazy"
            onLoad={() => setLoaded(true)}
            className={loaded ? 'is-loaded' : ''}
          />
        </div>
      </section>

      {/* modular & configurable */}
      <section className="sco-demo glass">
        <div className="sco-section-head">
          <h2><SlidersHorizontal size={20} color="var(--accent)" /> Modular &amp; configurable</h2>
          <a className="sco-open" href={CONFIG} target="_blank" rel="noreferrer">Open configurator <ExternalLink size={14} /></a>
        </div>
        <p className="sco-demo__hint">One standard engine, skinned and trimmed per retailer. Pick a tenant, recolor / rebrand, and toggle standard functions on/off — the kiosk, the flow and the payment options update live.</p>
        <div className="sco-twocol" style={{ marginBottom: '1rem' }}>
          <div className="sco-cap glass" style={{ '--c': 'var(--accent)' }}>
            <span className="sco-cap__icon"><Layers size={20} /></span>
            <h3>Modular</h3>
            <p>Switch any standard function off and the shopper flow gets shorter — each retailer ships only what they need.</p>
          </div>
          <div className="sco-cap glass" style={{ '--c': 'var(--accent-2)' }}>
            <span className="sco-cap__icon"><Palette size={20} /></span>
            <h3>Configurable</h3>
            <p>Recolor, re-image and rebrand every screen per tenant — Foodland today, Lotus&apos;s tomorrow — keeping the same standard functions.</p>
          </div>
        </div>
        <div className="sco-frame">
          {!cfgLoaded && <div className="sco-frame__loading">Loading configurator…</div>}
          <iframe
            title="SCO configurator"
            src={CONFIG}
            loading="lazy"
            onLoad={() => setCfgLoaded(true)}
            className={cfgLoaded ? 'is-loaded' : ''}
          />
        </div>
      </section>

      {/* shopper flow */}
      <section className="sco-block">
        <div className="sco-section-head"><h2>The shopper loop</h2></div>
        <div className="sco-flow">
          {FLOW.map((s, i) => {
            const { icon: Icon } = s
            return (
              <div key={s.n} className="sco-step">
                <span className="sco-step__icon"><Icon size={22} /></span>
                <span className="sco-step__n">{s.n}</span>
                <h3>{s.title}<small>{s.th}</small></h3>
                <p>{s.desc}</p>
                {i < FLOW.length - 1 && <span className="sco-step__arrow"><ArrowRight size={18} /></span>}
              </div>
            )
          })}
        </div>
      </section>

      {/* screens & overlays */}
      <section className="sco-block">
        <div className="sco-section-head"><h2>Every screen &amp; edge case</h2></div>
        <div className="sco-twocol">
          <div className="sco-list glass">
            <h3><LayoutGrid size={16} color="var(--accent)" /> Screens <span>{SCREENS.length}</span></h3>
            <div className="sco-tags">{SCREENS.map((s) => <span key={s} className="sco-tag">{s}</span>)}</div>
          </div>
          <div className="sco-list glass">
            <h3><CircleDot size={16} color="var(--accent-2)" /> Overlays / states <span>{OVERLAYS.length}</span></h3>
            <div className="sco-tags">{OVERLAYS.map((s) => <span key={s} className="sco-tag sco-tag--gold">{s}</span>)}</div>
          </div>
        </div>
      </section>

      {/* capabilities */}
      <section className="sco-block">
        <div className="sco-section-head"><h2>What the lane does</h2></div>
        <div className="sco-caps">
          {CAPS.map((c) => {
            const { icon: Icon } = c
            return (
              <div key={c.title} className="sco-cap glass" style={{ '--c': c.tone }}>
                <span className="sco-cap__icon"><Icon size={20} /></span>
                <h3>{c.title}</h3>
                <p>{c.text}</p>
              </div>
            )
          })}
        </div>
      </section>

      {/* promotions */}
      <section className="sco-block">
        <div className="sco-section-head"><h2><Tag size={18} color="var(--accent)" /> Promotions supported</h2></div>
        <div className="sco-twocol">
          <div className="sco-list glass">
            <h3 className="sco-list__good">Self-serve at the lane</h3>
            <div className="sco-tags">{PROMO_SELF.map((p) => <span key={p} className="sco-tag sco-tag--good">{p}</span>)}</div>
          </div>
          <div className="sco-list glass">
            <h3 className="sco-list__muted">Operation-controlled / API</h3>
            <div className="sco-tags">{PROMO_OPS.map((p) => <span key={p} className="sco-tag">{p}</span>)}</div>
          </div>
        </div>
        <p className="sco-fineprint">Codes from the Fujitsu BBY master in the kickoff — final scope &amp; naming to be confirmed with Foodland.</p>
      </section>

      {/* integration architecture */}
      <section className="sco-block">
        <div className="sco-section-head"><h2>How it connects</h2></div>
        <div className="sco-arch glass">
          <div className="sco-arch__chain">
            {ARCH.map((a, i) => {
              const { icon: Icon } = a
              return (
                <div key={a.label} className="sco-arch__node">
                  <span className="sco-arch__icon"><Icon size={20} /></span>
                  <span className="sco-arch__label">{a.label}</span>
                  {i < ARCH.length - 1 && <span className="sco-arch__link" aria-hidden="true" />}
                </div>
              )
            })}
          </div>
          <p className="sco-arch__note">
            Realtime down the chain. Cloud path: <strong>SCO → AWS API Gateway → Fujitsu API Engine</strong>
            (cache / DB / CICD) → Foodland loyalty, POS &amp; member servers. Fujitsu provides the lane &amp;
            API; <strong>Foodland keeps its systems of record.</strong>
          </p>
          <div className="sco-apis">
            <span className="sco-apis__label"><Plug size={14} /> API touch-points</span>
            {APIS.map((a) => (
              <span key={a.id} className="sco-api"><b>{a.id}</b> {a.label}</span>
            ))}
            <span className="sco-api"><b>EDC</b> payment interface</span>
          </div>
        </div>
      </section>

      {/* built on verified foundation */}
      <section className="sco-foundation glass">
        <div className="sco-foundation__text">
          <span className="sco-badge sco-badge--ghost">Evidence-driven</span>
          <h2>Built on a verified POS foundation</h2>
          <p>
            The data contract isn&apos;t guessed — we reverse-engineered the live POS (POSFront + ServiceTransfer)
            and the real <code>POSSDB</code> database. The SCO reads SKU / promo / tender / discount / points
            and writes <code>TPSTSal*</code> rows stamped <code>FTStaSentOnOff='0'</code> so the existing
            trigger + sync push every sale to HQ — and the Crystal reports keep working.
          </p>
          <div className="sco-foundation__links">
            <Link to="/flows" className="btn ghost"><Network size={15} /> System flows</Link>
            <Link to="/docs/data-dictionary" className="btn ghost"><Database size={15} /> Data dictionary</Link>
            <Link to="/methodology" className="btn ghost"><Sparkles size={15} /> The method</Link>
          </div>
        </div>
      </section>

      {/* scope & roadmap */}
      <section className="sco-block">
        <div className="sco-section-head"><h2>Prototype scope &amp; roadmap</h2></div>
        <div className="sco-twocol">
          <div className="sco-list glass">
            <h3 className="sco-list__muted"><Wrench size={15} /> Mocked today</h3>
            <ul className="sco-checks">{MOCKED.map((m) => <li key={m}>{m}</li>)}</ul>
          </div>
          <div className="sco-list glass">
            <h3 className="sco-list__good"><ListChecks size={15} /> Next for production</h3>
            <ul className="sco-checks sco-checks--good">{NEXT.map((m) => <li key={m}>{m}</li>)}</ul>
          </div>
        </div>
      </section>
    </div>
  )
}

export default SCO
