![Logo](./images/logo/vasst.png)

This repository is an historical record of work done for VASST, which sold video software plugins, ran video contests, and hosted a catalog of downloadable video recipes and templates that users could contribute to, rate and comment on.

- Created: **2002**
- Platform: ASP upgraded to VB.Net

# Innovations

The links below are to source files.

- Built an open-ended retail promotional system with [interfaces and base classes](/Code/entity/PromotionCriterion.vb) to allow arbitrary combinations of criteria and rewards (since the client kept dreaming up new ones). Criteria included things like sign-up date, particular item in cart and cart total over threshold.
- Built suite of custom HTML controls (and [setup IntelliSense for them](/Code/control/Intellisense.xsd)) that used reflection to read the criteria and other classes to generate friendly, easily managed forms.
- In order to support arbitrary criteria and rewards, built a [custom binary serialization solution](Code/data/File.vb) that could handle [deserializing old versions](Code/data/Migration.vb) of objects to instances having new fields (NoSQL solutions didn't exist yet).
- Used [graphics libraries](/Code/Draw.vb) to programatically enhance images and draw assets, like [ratings stars](/Code/control/Star.vb), on-the-fly (no usable `SVG` or `canvas` at the time).

