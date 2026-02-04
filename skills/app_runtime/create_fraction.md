# Skill: Create Fractions

## Description
This skill explains how to create and work with fractions in various programming contexts and mathematical applications. Fractions represent rational numbers as a ratio of two integers (numerator and denominator).

## Prerequisites
- Basic understanding of mathematical fractions
- Programming language of choice (C++, Python, JavaScript, etc.)

## Approaches

### 1) Using Built-in Fraction Libraries

#### Python
Python provides the `fractions` module for precise fraction arithmetic:

```python
from fractions import Fraction

# Create fractions in different ways
f1 = Fraction(3, 4)           # 3/4
f2 = Fraction(0.5)            # 1/2 (from decimal)
f3 = Fraction('1/3')          # 1/3 (from string)

# Arithmetic operations
result = f1 + f2              # 5/4
product = f1 * f3             # 1/4
```

#### JavaScript
```javascript
// Create a simple fraction object
class Fraction {
    constructor(numerator, denominator) {
        this.numerator = numerator;
        this.denominator = denominator;
        this.simplify();
    }
    
    simplify() {
        const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
        const divisor = gcd(this.numerator, this.denominator);
        this.numerator /= divisor;
        this.denominator /= divisor;
    }
}

const f = new Fraction(6, 8);  // Automatically simplifies to 3/4
```

#### C++
```cpp
#include <numeric>  // For std::gcd

class Fraction {
    int numerator;
    int denominator;
    
public:
    Fraction(int num, int den) : numerator(num), denominator(den) {
        if (den == 0) throw std::invalid_argument("Denominator cannot be zero");
        simplify();
    }
    
    void simplify() {
        int gcd = std::gcd(numerator, denominator);
        numerator /= gcd;
        denominator /= gcd;
    }
};
```

### 2) HTML/CSS for Visual Fractions

Display fractions visually on web pages:

```html
<!-- Using <sup> and <sub> tags -->
<p><sup>3</sup>&frasl;<sub>4</sub></p>

<!-- Using Unicode fraction characters -->
<p>¼ ½ ¾</p>

<!-- Using MathML -->
<math>
    <mfrac>
        <mn>3</mn>
        <mn>4</mn>
    </mfrac>
</math>
```

### 3) Key Operations

When implementing fractions, ensure these operations:

- **Addition**: Find common denominator, add numerators
- **Subtraction**: Find common denominator, subtract numerators
- **Multiplication**: Multiply numerators and denominators
- **Division**: Multiply by reciprocal
- **Simplification**: Divide both by GCD (Greatest Common Divisor)

## Best Practices

1. **Always simplify**: Reduce fractions to lowest terms
2. **Handle zero denominator**: Throw error or return invalid state
3. **Consider overflow**: Use appropriate integer types for large values
4. **Maintain immutability**: Return new fractions rather than modifying existing ones

## Done Criteria
You can successfully create, manipulate, and display fractions in your chosen environment, with proper error handling and simplified results.
