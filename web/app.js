"use strict";
const $ = id => document.getElementById(id);
let data = null, activeParams = null, currentPage = 1, busy = false;
const pageSize = 8;
const escapeHTML = value => String(value).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const ratingValue = b => b.rating ? Number(b.rating.split("/")[0]) : 0;
function message(text, error = false) { $("message").textContent = text; $("message").classList.toggle("error", error); }
function filtered() {
  const search = $("search").value.trim().toLowerCase(), rating = Number($("rating").value);
  return (data?.businesses || []).filter(b => (!search || [b.name,b.address,b.city_state,b.category].join(" ").toLowerCase().includes(search)) && (!rating || ratingValue(b) >= rating) && (!$("website").checked || b.website));
}
function renderRows() {
  const records = filtered(), pages = Math.max(1, Math.ceil(records.length / pageSize));
  currentPage = Math.min(currentPage, pages);
  const start = (currentPage - 1) * pageSize;
  $("rows").innerHTML = records.slice(start,start + pageSize).map(b => `<tr><td><strong>${escapeHTML(b.name)}</strong><small>${escapeHTML(b.address)}</small></td><td>${escapeHTML(b.phone)}<small>${escapeHTML(b.city_state)}</small></td><td>${b.rating ? `<span class="rating">★ ${ratingValue(b).toFixed(1)}</span>` : '<span class="missing">Unrated</span>'}</td><td>${b.website ? '<span class="site-yes">Available ↗</span>' : '<span class="missing">Not listed</span>'}</td></tr>`).join("") || '<tr><td colspan="4" class="empty-row">No matching businesses. Try another search or filter.</td></tr>';
  $("row-count").textContent = records.length ? `${start + 1}–${Math.min(start + pageSize,records.length)} of ${records.length} businesses` : "0 businesses";
  $("page-count").textContent = `${currentPage} / ${pages}`;
  $("previous").disabled = currentPage <= 1;
  $("next").disabled = currentPage >= pages;
  document.querySelectorAll(".export-button").forEach(b => b.disabled = !records.length || busy);
  $("export-rows").innerHTML = records.slice(0,8).map(b => `<tr><td>${escapeHTML(b.name)}</td><td>${escapeHTML(b.phone)}</td><td>${escapeHTML(b.address)}</td><td>${escapeHTML(b.city_state)}</td><td>${b.rating ? `${ratingValue(b).toFixed(1)} / 5` : "—"}</td><td>${escapeHTML(b.review_count || "—")}</td></tr>`).join("") || '<tr><td colspan="6" class="empty-row">No matching rows to export.</td></tr>';
  $("export-description").textContent = `${records.length} matching businesses`;
}
function render() {
  const records = data.businesses, stats = data.stats, rated = records.filter(b => b.rating), sites = records.filter(b => b.website).length;
  $("metric-total").textContent = records.length;
  $("metric-total-note").textContent = `${stats.duplicates} duplicate listings removed`;
  $("metric-pages").textContent = String(stats.pages).padStart(2,"0");
  $("metric-web").innerHTML = `${Math.round(sites / records.length * 100)}<em>%</em>`;
  $("metric-web-note").textContent = `${sites} of ${records.length} have a website`;
  $("metric-rating").innerHTML = `${rated.length ? (rated.reduce((sum,b) => sum + ratingValue(b),0) / rated.length).toFixed(1) : "—"} <em>/ 5</em>`;
  $("metric-rating-note").textContent = `Based on ${rated.length} rated businesses`;
  const category = activeParams.get("category");
  $("dataset-title").textContent = `${category[0].toUpperCase() + category.slice(1)} in ${activeParams.get("city")}`;
  $("duplicates").textContent = stats.duplicates;
  $("ads").textContent = stats.ads;
  const groups = [["5 ★",b=>ratingValue(b)===5],["4–4.5",b=>ratingValue(b)>=4 && ratingValue(b)<5],["3–3.5",b=>ratingValue(b)>=3 && ratingValue(b)<4],["Unrated",b=>!b.rating]];
  $("rating-bars").innerHTML = groups.map(([label,test])=>`<div class="bar-row"><span>${label}</span><div class="bar-track"><div class="bar" data-width="${records.filter(test).length / records.length * 100}"></div></div><b>${records.filter(test).length}</b></div>`).join("");
  // Width is a numeric presentation property, never sourced from scraped HTML.
  document.querySelectorAll(".bar").forEach(el => el.style.width = `${el.dataset.width}%`);
  $("activity-context").textContent = `${records.length} records · seed ${stats.seed}`;
  const events = [["Generated synthetic source pages",`${stats.pages} HTML pages prepared locally. No external requests.`],...stats.events.map(e=>[`Parsed page ${String(e.page).padStart(2,"0")}`,`${e.parsed} organic listings found · ${e.added} new businesses retained.`]),["Normalized the dataset",`${stats.duplicates} repeated listings removed. ${stats.ads} sponsored cards excluded.`],["Ready to export",`${records.length} unique records available as a formatted Excel workbook.`]];
  $("events").innerHTML = events.map(([title,detail],i)=>`<li><span class="step">${i+1}</span><div><b>${escapeHTML(title)}</b><p>${escapeHTML(detail)}</p></div></li>`).join("");
  renderRows();
}
async function runDemo(event) {
  event?.preventDefault();
  if (busy) return;
  busy = true; $("run").disabled = true; $("run").textContent = "Processing…";
  document.querySelectorAll(".export-button").forEach(b=>b.disabled=true);
  message("Parsing generated directory pages…");
  const params = new URLSearchParams({category:$("category").value, city:$("city").value, count:$("count").value, seed:"42"});
  try {
    const response = await fetch(`/api/demo?${params}`);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Unable to run the demo.");
    data = payload; activeParams = params; currentPage = 1;
    $("search").value = ""; $("rating").value = "0"; $("website").checked = false;
    message(""); render();
  } catch (error) { message(error.message,true); }
  finally { busy = false; $("run").disabled = false; $("run").innerHTML = 'Run pipeline <span>→</span>'; if(data) renderRows(); }
}
async function exportRows() {
  if (!data || busy || !filtered().length) return;
  const params = new URLSearchParams(activeParams);
  params.set("search",$("search").value); params.set("rating",$("rating").value); params.set("website",$("website").checked);
  document.querySelectorAll(".export-button").forEach(b=>b.disabled=true);
  try {
    const response = await fetch(`/api/export?${params}`);
    if (!response.ok) { const payload = await response.json(); throw new Error(payload.error || "Export failed."); }
    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a"); link.href = url; link.download = `${activeParams.get("category")}_demo.xlsx`; link.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
    message(`Excel export created with ${filtered().length} matching businesses.`);
  } catch(error) { message(error.message,true); }
  finally { renderRows(); }
}
document.querySelectorAll(".nav").forEach(button=>button.addEventListener("click",()=>{
  document.querySelectorAll(".nav").forEach(b=>b.classList.toggle("active",b===button));
  document.querySelectorAll(".view").forEach(v=>v.hidden=v.id!==`${button.dataset.view}-view`);
  $("view-label").textContent = button.textContent.trim().slice(1).trim(); message("");
}));
$("run-form").addEventListener("submit",runDemo);
["search","rating","website"].forEach(id=>$(id).addEventListener("input",()=>{currentPage=1;renderRows();}));
$("previous").addEventListener("click",()=>{currentPage--;renderRows();});
$("next").addEventListener("click",()=>{currentPage++;renderRows();});
document.querySelectorAll(".export-button").forEach(button=>button.addEventListener("click",exportRows));
runDemo();
