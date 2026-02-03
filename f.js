/**
 * @param {number[][]} people
 * @return {number[][]}
 */
var reconstructQueue = function(people) {
    // select zeros and sort by first number
    // record first number to new array
    // go from shortest person in remaining and check that second number
    // is equal to number of people heigher in new array

    let arrHeights = [];
    let arrMins = [];
    let ongingMin = -1;
    for (let i = 0; i < people.length; i++) {
        let min = i;


        for (let j = i; j < people.length; j++) {

            const currentHeight = people[j][0];
            let inFrontSmol = 0;
            for (let k = 0; k < arrHeights.length; k++) {
                if (arrHeights[k] >= currentHeight)
                    inFrontSmol++;
            }


            if (people[j][1] == inFrontSmol) {
                if (ongingMin == -1) {
                    min = j;
                    ongingMin = people[min][0];
                } else if (people[j][0] < ongingMin) {
                    min = j;
                    ongingMin = people[min][0];
                }
            }


        }

        [people[i], people[min]] = [people[min], people[i]];
        arrHeights.push(people[i][0]);
        ongingMin = -1;
    }
    return people;
};


people = [[9,0],[7,0],[1,9],[3,0],[2,7],[5,3],[6,0],[3,4],[6,2],[5,2]]

console.log(reconstructQueue(people));
let i = 3;