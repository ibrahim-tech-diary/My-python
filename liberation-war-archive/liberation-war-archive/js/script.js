// Document search
function searchDocument(){

let input = document.getElementById("search").value.toLowerCase();

let cards = document.getElementsByClassName("doc-card");

for(let i=0;i<cards.length;i++){

let text = cards[i].innerText.toLowerCase();

if(text.includes(input)){
cards[i].style.display="block";
}
else{
cards[i].style.display="none";
}

}

}